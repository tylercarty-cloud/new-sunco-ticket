(function () {
  'use strict';

  var client = ZAFClient.init();
  var SUNCO_CONVERSATION_FIELD_ID = '46936995990164';
  var TICKET_TITLE_FIELD_ID = '46410167947924';
  var REQUEST_TYPE_FIELD_ID = '42571762576276';
  var AUTO_TAGS = ['mccomplete', 'manual_ticket'];
  var REQUEST_TIMEOUT_MS = 30000;
  var TAG_POLL_INTERVAL_MS = 500;
  var TAG_POLL_MAX_ATTEMPTS = 60;

  var state = {
    requesterName: null,
    requesterExternalId: null,
    requesterId: null,
    suncoAppId: null,
    suncoKeyId: null,
    suncoKeySecret: null,
    originalCustomFields: [],
    originalTicketId: null,
    isReady: false
  };

  var requesterDisplay = document.getElementById('requesterDisplay');
  var externalIdDisplay = document.getElementById('externalIdDisplay');
  var ticketForm = document.getElementById('ticketForm');
  var ticketSubject = document.getElementById('ticketSubject');
  var ticketTitle = document.getElementById('ticketTitle');
  var requestType = document.getElementById('requestType');
  var ticketMessage = document.getElementById('ticketMessage');
  var createBtn = document.getElementById('createBtn');
  var progressSteps = document.getElementById('progressSteps');
  var statusBar = document.getElementById('statusBar');

  function log(label, data) {
    console.log('[SunCo App] ' + label, data !== undefined ? data : '');
  }

  function logError(label, err) {
    console.error('[SunCo App] ' + label, err);
    if (err) {
      console.error('[SunCo App] Error details:', JSON.stringify({
        message: err.message, status: err.status,
        statusText: err.statusText, responseText: err.responseText,
        responseJSON: err.responseJSON
      }, null, 2));
    }
  }

  // ── Initialization ──

  client.on('app.registered', function () {
    log('App registered');
    resizeApp();
    loadSettings();
    loadRequesterInfo();
    loadRequestTypeOptions();
    loadOriginalCustomFields();
  });

  function resizeApp() {
    setTimeout(function () {
      var height = document.body.scrollHeight + 16;
      client.invoke('resize', { width: '100%', height: height + 'px' });
    }, 50);
  }

  client.on('ticket.requester.id.changed', function () {
    log('Requester changed');
    loadRequesterInfo();
  });

  function loadSettings() {
    client.metadata().then(function (metadata) {
      state.suncoAppId = metadata.settings.suncoAppId;
      state.suncoKeyId = metadata.settings.suncoKeyId;
      state.suncoKeySecret = metadata.settings.suncoKeySecret;
      log('Settings loaded', {
        suncoAppId: state.suncoAppId || '(NOT SET)',
        suncoKeyId: state.suncoKeyId ? '(set)' : '(NOT SET)',
        suncoKeySecret: state.suncoKeySecret ? '(set)' : '(NOT SET)'
      });
    }).catch(function (err) {
      logError('Failed to load settings', err);
    });
  }

  function loadRequesterInfo() {
    client.get([
      'ticket.requester.name',
      'ticket.requester.externalId',
      'ticket.requester.id'
    ]).then(function (data) {
      state.requesterName = data['ticket.requester.name'];
      state.requesterExternalId = data['ticket.requester.externalId'];
      state.requesterId = data['ticket.requester.id'];

      log('Requester info loaded', {
        name: state.requesterName,
        externalId: state.requesterExternalId,
        requesterId: state.requesterId
      });

      requesterDisplay.textContent = state.requesterName || 'Unknown';
      requesterDisplay.classList.remove('missing');

      if (state.requesterExternalId) {
        externalIdDisplay.textContent = state.requesterExternalId;
        externalIdDisplay.classList.remove('missing');
        state.isReady = true;
        createBtn.disabled = false;
      } else {
        externalIdDisplay.textContent = 'No external ID found for this requester';
        externalIdDisplay.classList.add('missing');
        state.isReady = false;
        createBtn.disabled = true;
      }
    }).catch(function (err) {
      logError('Failed to load requester info', err);
      requesterDisplay.textContent = 'Unable to load requester';
      requesterDisplay.classList.add('missing');
      externalIdDisplay.textContent = 'Unable to load external ID';
      externalIdDisplay.classList.add('missing');
      state.isReady = false;
      createBtn.disabled = true;
    });
  }

  function loadRequestTypeOptions() {
    client.request({
      url: '/api/v2/ticket_fields/' + REQUEST_TYPE_FIELD_ID + '.json',
      type: 'GET'
    }).then(function (response) {
      var data = (typeof response === 'string') ? JSON.parse(response) : response;
      if (data.ticket_field && data.ticket_field.custom_field_options) {
        var options = data.ticket_field.custom_field_options;
        log('Request Type options loaded', options.length + ' options');

        requestType.innerHTML = '<option value="">-- Select --</option>';
        options.forEach(function (opt) {
          var el = document.createElement('option');
          el.value = opt.value;
          el.textContent = opt.name;
          requestType.appendChild(el);
        });
      }
    }).catch(function (err) {
      logError('Failed to load Request Type options', err);
    });
  }

  function loadOriginalCustomFields() {
    client.get('ticket.id').then(function (data) {
      var ticketId = data['ticket.id'];
      if (!ticketId) {
        log('No ticket ID yet, skipping custom field load');
        return;
      }

      state.originalTicketId = ticketId;
      log('Original ticket ID stored:', state.originalTicketId);

      return client.request({
        url: '/api/v2/tickets/' + ticketId + '.json',
        type: 'GET'
      }).then(function (response) {
        var data = (typeof response === 'string') ? JSON.parse(response) : response;
        if (data.ticket && data.ticket.custom_fields) {
          state.originalCustomFields = data.ticket.custom_fields.filter(function (f) {
            return f.value !== null && f.value !== '';
          });
          log('Original custom fields loaded', state.originalCustomFields.length + ' fields with values');
        }
      });
    }).catch(function (err) {
      logError('Failed to load original custom fields', err);
    });
  }

  // ── Form submission ──

  ticketForm.addEventListener('submit', function (e) {
    e.preventDefault();
    if (!state.isReady) {
      log('Form submitted but not ready', state);
      return;
    }

    var subject = ticketSubject.value.trim();
    var ticketTitleVal = ticketTitle.value.trim();
    var requestTypeVal = requestType.value;
    var message = ticketMessage.value.trim();
    if (!subject || !message) return;

    log('Starting ticket creation', {
      subject: subject, ticketTitle: ticketTitleVal,
      requestType: requestTypeVal, message: message
    });
    createSunCoTicket(subject, message, ticketTitleVal, requestTypeVal);
  });

  // ── Main workflow ──
  //
  // 1. Create SunCo conversation (proper requester/channel linkage).
  // 2. Send user message — this triggers the SunCo integration to create a
  //    Zendesk ticket.  Immediately begin aggressive polling (500ms) so we
  //    find the ticket as fast as possible.
  // 3. The instant the ticket is found, fire a tag-only PUT to stamp
  //    mccomplete + manual_ticket before the response bot can act.
  // 4. Then do the full update (subject, requester, custom fields).

  function createSunCoTicket(subject, message, ticketTitleVal, requestTypeVal) {
    try {
      setFormEnabled(false);
      showProgress();
      hideStatus();

      var conversationId;
      var workflowStartedAt = new Date().toISOString();

      log('Step 1: Creating SunCo conversation for externalId:', state.requesterExternalId);
      log('Workflow started at:', workflowStartedAt);
      setStepState(1, 'active');

      createSunCoConversation(state.requesterExternalId)
        .then(function (convId) {
          conversationId = convId;
          log('Step 1 complete. Conversation ID:', conversationId);
          setStepState(1, 'done');
          setStepState(2, 'active');
          log('Step 2: Sending user message to trigger ticket creation');
          return sendSunCoMessage(conversationId, message, 'user');
        })
        .then(function () {
          log('Step 2 complete. User message sent.');
          setStepState(2, 'done');
          setStepState(3, 'active');
          log('Step 3: Polling for ticket and applying tags immediately');
          return pollAndTagTicket(workflowStartedAt);
        })
        .then(function (ticketId) {
          log('Step 3 complete. Ticket #' + ticketId + ' found and tagged.');
          setStepState(3, 'done');

          if (ticketId) {
            setStepState(4, 'active');
            log('Step 4: Updating ticket with subject, custom fields, and conversation ID');
            return updateTicketFields(ticketId, subject, ticketTitleVal, requestTypeVal, conversationId).then(function () {
              log('Step 4 complete.');
              setStepState(4, 'done');
              return ticketId;
            });
          }
          return ticketId;
        })
        .then(function (ticketId) {
          if (ticketId) {
            showStatus(
              'success',
              'Ticket <a href="#" id="openTicketLink">#' + ticketId + '</a> created via Sunshine Conversations.'
            );
            var link = document.getElementById('openTicketLink');
            if (link) {
              link.addEventListener('click', function (ev) {
                ev.preventDefault();
                client.invoke('routeTo', 'ticket', ticketId);
              });
            }
          } else {
            showStatus(
              'success',
              'Conversation created and message sent. The ticket should appear shortly in your queue.'
            );
          }

          ticketSubject.value = '';
          ticketTitle.value = '';
          requestType.value = '';
          ticketMessage.value = '';
          setFormEnabled(true);
        })
        .catch(function (err) {
          logError('Workflow failed', err);
          var failedStep = getCurrentActiveStep();
          if (failedStep) setStepState(failedStep, 'failed');
          showStatus('error', extractError(err) + '<br><br><small>Check browser console (F12) for details.</small>');
          setFormEnabled(true);
        });
    } catch (syncError) {
      logError('Synchronous error in createSunCoTicket', syncError);
      showStatus('error', 'JavaScript error: ' + syncError.message);
      setFormEnabled(true);
    }
  }

  // ── SunCo API ──

  function makeSunCoRequest(path, body, method) {
    var httpMethod = method || 'POST';
    var url = 'https://impiricussupport.zendesk.com/sc' + path;
    var authToken = 'Basic ' + btoa(state.suncoKeyId + ':' + state.suncoKeySecret);

    log('SunCo API Request', JSON.stringify({ url: url, method: httpMethod, body: body }, null, 2));

    var reqOpts = {
      url: url,
      type: httpMethod,
      contentType: 'application/json',
      cors: true,
      headers: { Authorization: authToken }
    };
    if (body) {
      reqOpts.data = JSON.stringify(body);
    }

    return withTimeout(
      client.request(reqOpts).then(function (response) {
        log('SunCo API Response', response);
        return response;
      }).catch(function (err) {
        logError('SunCo API Error', err);
        throw err;
      }),
      REQUEST_TIMEOUT_MS,
      'SunCo API request timed out.'
    );
  }

  function createSunCoConversation(externalId) {
    var path = '/v2/apps/' + state.suncoAppId + '/conversations';
    return makeSunCoRequest(path, {
      type: 'personal',
      participants: [{ userExternalId: externalId }],
      metadata: { stop_ai_escalated: true }
    }).then(function (response) {
      var data = (typeof response === 'string') ? JSON.parse(response) : response;
      log('Parsed conversation response', data);
      if (data.conversation && data.conversation.id) {
        return data.conversation.id;
      }
      throw new Error('No conversation ID in response: ' + JSON.stringify(data));
    });
  }

  function sendSunCoMessage(conversationId, messageText, authorType) {
    var path = '/v2/apps/' + state.suncoAppId + '/conversations/' + conversationId + '/messages';
    var author = { type: authorType };
    if (authorType === 'user') {
      author.userExternalId = state.requesterExternalId;
    }
    return makeSunCoRequest(path, {
      author: author,
      content: { type: 'text', text: messageText }
    });
  }

  // ── Zendesk API ──

  function applyTagsToTicket(ticketId) {
    log('Applying tags to ticket #' + ticketId + ':', AUTO_TAGS);
    return client.request({
      url: '/api/v2/tickets/' + ticketId + '.json',
      type: 'PUT',
      contentType: 'application/json',
      data: JSON.stringify({ ticket: { tags: AUTO_TAGS } })
    }).then(function (response) {
      log('Tags applied to ticket #' + ticketId);
      return response;
    }).catch(function (err) {
      logError('Failed to apply tags to ticket #' + ticketId, err);
      throw err;
    });
  }

  function updateTicketFields(ticketId, subject, ticketTitleVal, requestTypeVal, conversationId) {
    var customFields = buildCustomFields(ticketTitleVal, requestTypeVal, conversationId);

    var updateData = {
      ticket: {
        requester_id: state.requesterId,
        subject: subject,
        tags: AUTO_TAGS,
        custom_fields: customFields
      }
    };

    log('Updating ticket #' + ticketId + ' with payload:', JSON.stringify(updateData, null, 2));

    return client.request({
      url: '/api/v2/tickets/' + ticketId + '.json',
      type: 'PUT',
      contentType: 'application/json',
      data: JSON.stringify(updateData)
    }).then(function (response) {
      log('Ticket updated successfully', response);
      return response;
    }).catch(function (err) {
      logError('Failed to update ticket', err);
      throw err;
    });
  }

  function buildCustomFields(ticketTitleVal, requestTypeVal, conversationId) {
    var fieldsMap = {};

    state.originalCustomFields.forEach(function (f) {
      fieldsMap[String(f.id)] = f.value;
    });

    if (ticketTitleVal) {
      fieldsMap[TICKET_TITLE_FIELD_ID] = ticketTitleVal;
    }
    if (requestTypeVal) {
      fieldsMap[REQUEST_TYPE_FIELD_ID] = requestTypeVal;
    }
    if (conversationId) {
      fieldsMap[SUNCO_CONVERSATION_FIELD_ID] = conversationId;
    }

    var result = [];
    Object.keys(fieldsMap).forEach(function (id) {
      result.push({ id: Number(id), value: fieldsMap[id] });
    });

    log('Built custom fields (' + result.length + ' fields)', JSON.stringify(result, null, 2));
    return result;
  }

  // ── Polling ──
  //
  // Aggressive polling: start immediately after the message is sent, check
  // every 500ms.  The moment a new ticket is found, fire a tag-only PUT
  // before returning the ticket ID so tags land as fast as possible.

  function pollAndTagTicket(startedAt) {
    log('Aggressive poll: looking for ticket by requester ' + state.requesterId + ' created after ' + startedAt);
    var attempts = 0;

    return new Promise(function (resolve, reject) {
      function check() {
        attempts++;
        log('Tag-poll attempt ' + attempts + ' of ' + TAG_POLL_MAX_ATTEMPTS);

        findNewestTicket(startedAt)
          .then(function (ticketId) {
            if (ticketId) {
              log('Found ticket #' + ticketId + ' on attempt ' + attempts + ', applying tags now');
              return applyTagsToTicket(ticketId).then(function () {
                resolve(ticketId);
              });
            }
            if (attempts >= TAG_POLL_MAX_ATTEMPTS) {
              log('Max poll attempts reached, no ticket found');
              resolve(null);
              return;
            }
            setTimeout(check, TAG_POLL_INTERVAL_MS);
          })
          .catch(function (err) {
            logError('Error during tag-poll', err);
            if (attempts >= TAG_POLL_MAX_ATTEMPTS) {
              resolve(null);
            } else {
              setTimeout(check, TAG_POLL_INTERVAL_MS);
            }
          });
      }

      check();
    });
  }

  function isNewTicket(ticket, startedAt) {
    if (!ticket || !ticket.created_at) return false;
    if (ticket.created_at < startedAt) return false;
    if (state.originalTicketId && String(ticket.id) === String(state.originalTicketId)) return false;
    return true;
  }

  function findNewestTicket(startedAt) {
    var cacheBust = '_=' + Date.now();

    var viaList = client.request({
      url: '/api/v2/users/' + state.requesterId + '/tickets/requested.json?sort_by=created_at&sort_order=desc&per_page=5&' + cacheBust,
      type: 'GET'
    }).then(function (response) {
      var data = (typeof response === 'string') ? JSON.parse(response) : response;
      var tickets = data.tickets || [];
      log('[List] Returned ' + tickets.length + ' ticket(s)');
      for (var i = 0; i < tickets.length; i++) {
        var t = tickets[i];
        log('[List] #' + t.id + ' created ' + t.created_at);
        if (isNewTicket(t, startedAt)) {
          return t.id;
        }
      }
      return null;
    }).catch(function (err) {
      logError('List endpoint failed', err);
      return null;
    });

    var searchQuery = 'type:ticket requester_id:' + state.requesterId + ' created>' + startedAt;
    var viaSearch = client.request({
      url: '/api/v2/search.json?query=' + encodeURIComponent(searchQuery) + '&sort_by=created_at&sort_order=desc&' + cacheBust,
      type: 'GET'
    }).then(function (response) {
      var data = (typeof response === 'string') ? JSON.parse(response) : response;
      var results = data.results || [];
      log('[Search] Returned ' + results.length + ' result(s)');
      for (var i = 0; i < results.length; i++) {
        var t = results[i];
        log('[Search] #' + t.id + ' created ' + t.created_at);
        if (isNewTicket(t, startedAt)) {
          return t.id;
        }
      }
      return null;
    }).catch(function (err) {
      logError('Search endpoint failed', err);
      return null;
    });

    return Promise.all([viaList, viaSearch]).then(function (results) {
      return results[0] || results[1] || null;
    });
  }

  // ── Utilities ──

  function withTimeout(promise, ms, message) {
    return new Promise(function (resolve, reject) {
      var timedOut = false;
      var timer = setTimeout(function () {
        timedOut = true;
        reject(new Error(message));
      }, ms);
      promise.then(function (result) {
        if (!timedOut) { clearTimeout(timer); resolve(result); }
      }).catch(function (err) {
        if (!timedOut) { clearTimeout(timer); reject(err); }
      });
    });
  }

  function extractError(err) {
    var parts = [];
    if (err && err.status) parts.push('HTTP ' + err.status + (err.statusText ? ' ' + err.statusText : ''));
    if (err && err.responseJSON) {
      var json = err.responseJSON;
      if (json.error) {
        if (typeof json.error === 'string') parts.push(json.error);
        else if (json.error.description) parts.push(json.error.description);
        else if (json.error.message) parts.push(json.error.message);
        else parts.push(JSON.stringify(json.error));
      } else if (json.message) parts.push(json.message);
      else parts.push(JSON.stringify(json));
    } else if (err && err.responseText) {
      parts.push(err.responseText.substring(0, 500));
    } else if (err && err.message) {
      parts.push(err.message);
    } else if (typeof err === 'string') {
      parts.push(err);
    }
    return parts.length > 0 ? parts.join(' - ') : 'Unknown error occurred';
  }

  // ── UI Helpers ──

  function setFormEnabled(enabled) {
    ticketSubject.disabled = !enabled;
    ticketTitle.disabled = !enabled;
    requestType.disabled = !enabled;
    ticketMessage.disabled = !enabled;
    createBtn.disabled = !enabled;
    createBtn.innerHTML = enabled ? 'Create SunCo Ticket' : '<span class="spinner"></span>Creating...';
  }

  function setStepState(stepNum, stepState) {
    var step = document.getElementById('step' + stepNum);
    if (!step) return;
    step.setAttribute('class', 'step ' + stepState);
    var icon = step.querySelector('.step-icon');
    if (icon) {
      icon.setAttribute('class', 'step-icon ' + stepState);
      if (stepState === 'done') {
        icon.innerHTML = '<path d="M8 1a7 7 0 100 14A7 7 0 008 1zm3.7 5.3l-4 4a1 1 0 01-1.4 0l-2-2a1 1 0 111.4-1.4L7 8.2l3.3-3.3a1 1 0 011.4 1.4z"/>';
      } else if (stepState === 'failed') {
        icon.innerHTML = '<path d="M8 1a7 7 0 100 14A7 7 0 008 1zm2.8 8.4a1 1 0 11-1.4 1.4L8 9.4l-1.4 1.4a1 1 0 11-1.4-1.4L6.6 8 5.2 6.6a1 1 0 111.4-1.4L8 6.6l1.4-1.4a1 1 0 111.4 1.4L9.4 8l1.4 1.4z"/>';
      } else {
        icon.innerHTML = '<circle cx="8" cy="8" r="6"/>';
      }
    }
  }

  function getCurrentActiveStep() {
    for (var i = 1; i <= 4; i++) {
      var step = document.getElementById('step' + i);
      if (step && step.getAttribute('class').indexOf('active') !== -1) return i;
    }
    return null;
  }

  function showProgress() {
    progressSteps.classList.add('visible');
    setStepState(1, 'pending');
    setStepState(2, 'pending');
    setStepState(3, 'pending');
    setStepState(4, 'pending');
    resizeApp();
  }

  function showStatus(type, message) {
    statusBar.className = 'status-bar ' + type;
    statusBar.innerHTML = message;
    resizeApp();
  }

  function hideStatus() {
    statusBar.className = 'status-bar';
    statusBar.innerHTML = '';
  }

})();
