$(function() {
  // App configuration
  var authEndpoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?';
  var apiEndpoint = 'https://outlook.office.com/api/v2.0';
  var redirectUri = 'http://localhost:8080';
  var appId = 'YOUR APP ID HERE';
  var scopes = 'openid profile https://outlook.office.com/mail.read https://outlook.office.com/calendars.read https://outlook.office.com/contacts.read';

  // Check for browser support for sessionStorage
  if (typeof(Storage) === 'undefined') {
    render('#unsupportedbrowser');
    return;
  }
  
  render(window.location.hash);

  $(window).on('hashchange', function() {
    render(window.location.hash);
  });
  
  function render(hash) {

    var action = hash.split('=')[0];

    // Hide everything
    $('.main-container .page').hide();

    // Check for presence of access token
    var isAuthenticated = (sessionStorage.accessToken != null && sessionStorage.accessToken.length > 0);
    renderNav(isAuthenticated);
    renderTokens();
    
    var pagemap = {
      
      // Welcome page
      '': function() {
        renderWelcome(isAuthenticated);
      },

      // Receive access token
      '#access_token': function() {
        handleTokenResponse(hash);             
      },

      // Signout
      '#signout': function () {
        clearUserState();
        
        // Redirect to home page
        window.location.hash = '#';
      },

      // Error display
      '#error': function () {
        var errorresponse = parseHashParams(hash);
        renderError(errorresponse.error, errorresponse.error_description);
      },

      // Display inbox
      '#inbox': function () {
        if (isAuthenticated) {
          renderInbox();  
        } else {
          // Redirect to home page
          window.location.hash = '#';
        }
      },

      // Display calendar
      '#calendar': function () {
        if (isAuthenticated) {
          renderCalendar();  
        } else {
          // Redirect to home page
          window.location.hash = '#';
        }
      },

      // Display contacts
      '#contacts': function () {
        if (isAuthenticated) {
          renderContacts();  
        } else {
          // Redirect to home page
          window.location.hash = '#';
        }
      },

      // Shown if browser doesn't support session storage
      '#unsupportedbrowser': function () {
        $('#unsupported').show();
      }
    }
    
    if (pagemap[action]){
      pagemap[action]();
    } else {
      // Redirect to home page
      window.location.hash = '#';
    }
  }

  function setActiveNav(navId) {
    $('#navbar').find('li').removeClass('active');
    $(navId).addClass('active');
  }

  function renderNav(isAuthed) {
    if (isAuthed) {
      $('.authed-nav').show();
    } else {
      $('.authed-nav').hide();
    }
  }

  function renderTokens() {
    if (sessionStorage.accessToken) {
      // For demo purposes display the token and expiration
      var expireDate = new Date(parseInt(sessionStorage.tokenExpires));
      $('#token', window.parent.document).text(sessionStorage.accessToken);
      $('#expires-display', window.parent.document).text(expireDate.toLocaleDateString() + ' ' + expireDate.toLocaleTimeString());
      if (sessionStorage.idToken) {
        $('#id-token', window.parent.document).text(sessionStorage.idToken);
      }
      $('#token-display', window.parent.document).show();
    } else {
      $('#token-display', window.parent.document).hide();
    }
  }

  function renderError(error, description) {
    $('#error-name', window.parent.document).text('An error occurred: ' + decodePlusEscaped(error));
    $('#error-desc', window.parent.document).text(decodePlusEscaped(description));
    $('#error-display', window.parent.document).show();
  }
  
  function renderWelcome(isAuthed) {
    setActiveNav('#home-nav');
    if (isAuthed) {
      $('#username').text(sessionStorage.userDisplayName);
      $('#logged-in-welcome').show();
    } else {
      $('#connect-button').attr('href', buildAuthUrl());
      $('#signin-prompt').show();
    }
  }

  function renderInbox() {
    setActiveNav('#inbox-nav');
    $('#inbox-status').text('Loading...');
    $('#message-list').empty();
    $('#inbox').show();
    // Get user's email address
    getUserEmailAddress(function(userEmail, error) {
      if (error) {
        renderError('getUserEmailAddress failed', error.responseText);
      } else {
        getUserInboxMessages(userEmail, function(messages, error){
          $('#inbox-status').text('Here are the 10 most recent messages in your inbox.');
          var templateSource = $('#msg-list-template').html();
          var template = Handlebars.compile(templateSource);

          var msgList = template({messages: messages});
          $('#message-list').append(msgList);
        });
      }
    });
  }

  function renderCalendar() {
    setActiveNav('#calendar-nav');
    $('#calendar-status').text('Loading...');
    $('#event-list').empty();
    $('#calendar').show();
    // Get user's email address
    getUserEmailAddress(function(userEmail, error) {
      if (error) {
        renderError('getUserEmailAddress failed', error.responseText);
      } else {
        getUserEvents(userEmail, function(events, error){
          $('#calendar-status').text('Here are the 10 most recently created events on your calendar.');
          var templateSource = $('#event-list-template').html();
          var template = Handlebars.compile(templateSource);

          var eventList = template({events: events});
          $('#event-list').append(eventList);
        });
      }
    });
  }

  function renderContacts() {
    setActiveNav('#contacts-nav');
    $('#contacts-status').text('Loading...');
    $('#contact-list').empty();
    $('#contacts').show();
    // Get user's email address
    getUserEmailAddress(function(userEmail, error) {
      if (error) {
        renderError('getUserEmailAddress failed', error.responseText);
      } else {
        getUserContacts(userEmail, function(contacts, error){
          $('#contacts-status').text('Here are your first 10 contacts.');
          var templateSource = $('#contact-list-template').html();
          var template = Handlebars.compile(templateSource);

          var contactList = template({contacts: contacts});
          $('#contact-list').append(contactList);
        });
      }
    });
  }

  // OAUTH FUNCTIONS =============================

  function buildAuthUrl() {
    // Generate random values for state and nonce
    sessionStorage.authState = guid();
    sessionStorage.authNonce = guid();

    var authParams = {
      response_type: 'id_token token',
      client_id: appId,
      redirect_uri: redirectUri,
      scope: scopes,
      state: sessionStorage.authState,
      nonce: sessionStorage.authNonce,
      response_mode: 'fragment'
    };
    
    return authEndpoint + $.param(authParams);
  }

  function handleTokenResponse(hash) {
    // If this was a silent request remove the iframe
    $('#auth-iframe').remove();

    // clear tokens
    sessionStorage.removeItem('accessToken');
    sessionStorage.removeItem('idToken');

    var tokenresponse = parseHashParams(hash);

    // Check that state is what we sent in sign in request
    if (tokenresponse.state != sessionStorage.authState) {
      sessionStorage.removeItem('authState');
      sessionStorage.removeItem('authNonce');
      // Report error
      window.location.hash = '#error=Invalid+state&error_description=The+state+in+the+authorization+response+did+not+match+the+expected+value.+Please+try+signing+in+again.';
      return;
    }

    sessionStorage.authState = '';
    sessionStorage.accessToken = tokenresponse.access_token;
    
    // Get the number of seconds the token is valid for,
    // Subract 5 minutes (300 sec) to account for differences in clock settings
    // Convert to milliseconds
    var expiresin = (parseInt(tokenresponse.expires_in) - 300) * 1000;
    var now = new Date();
    var expireDate = new Date(now.getTime() + expiresin);
    sessionStorage.tokenExpires = expireDate.getTime();

    sessionStorage.idToken = tokenresponse.id_token;

    validateIdToken(function(isValid) {
      if (isValid) {
        // Re-render token to handle refresh
        renderTokens();
        
        // Redirect to home page
        window.location.hash = '#';
      } else {
        clearUserState();
        // Report error
        window.location.hash = '#error=Invalid+ID+token&error_description=ID+token+failed+validation,+please+try+signing+in+again.';
      }
    });
  }

  function validateIdToken(callback) {
    // Per Azure docs (and OpenID spec), we MUST validate
    // the ID token before using it. However, full validation
    // of the signature currently requires a server-side component
    // to fetch the public signing keys from Azure. This sample will
    // skip that part (technically violating the OpenID spec) and do
    // minimal validation

    if (null == sessionStorage.idToken || sessionStorage.idToken.length <= 0) {
      callback(false);
    }

    // JWT is in three parts seperated by '.'
    var tokenParts = sessionStorage.idToken.split('.');
    if (tokenParts.length != 3){
      callback(false);
    }

    // Parse the token parts
    var header = KJUR.jws.JWS.readSafeJSONString(b64utoutf8(tokenParts[0]));
    var payload = KJUR.jws.JWS.readSafeJSONString(b64utoutf8(tokenParts[1]));

    // Check the nonce
    if (payload.nonce != sessionStorage.authNonce) {
      sessionStorage.authNonce = '';
      callback(false);
    }

    sessionStorage.authNonce = '';

    // Check the audience
    if (payload.aud != appId) {
      callback(false);
    }

    // Check the issuer
    // Should be https://login.microsoftonline.com/{tenantid}/v2.0
    if (payload.iss !== 'https://login.microsoftonline.com/' + payload.tid + '/v2.0') {
      callback(false);
    }

    // Check the valid dates
    var now = new Date();
    // To allow for slight inconsistencies in system clocks, adjust by 5 minutes
    var notBefore = new Date((payload.nbf - 300) * 1000);
    var expires = new Date((payload.exp + 300) * 1000);
    if (now < notBefore || now > expires) {
      callback(false);
    }

    // Now that we've passed our checks, save the bits of data
    // we need from the token.

    sessionStorage.userDisplayName = payload.name;
    sessionStorage.userSigninName = payload.preferred_username;

    // Per the docs at:
    // https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-protocols-implicit/#send-the-sign-in-request
    // Check if this is a consumer account so we can set domain_hint properly
    sessionStorage.userDomainType = 
      payload.tid === '9188040d-6c67-4c5b-b112-36a304b66dad' ? 'consumers' : 'organizations';

    callback(true);
  }

  function makeSilentTokenRequest(callback) {
    // Build up a hidden iframe
    var iframe = $('<iframe/>');
    iframe.attr('id', 'auth-iframe');
    iframe.attr('name', 'auth-iframe');
    iframe.appendTo('body');
    iframe.hide();

    iframe.load(function() {
      callback(sessionStorage.accessToken);
    });
    
    iframe.attr('src', buildAuthUrl() + '&prompt=none&domain_hint=' + 
      sessionStorage.userDomainType + '&login_hint=' + 
      sessionStorage.userSigninName);
  }

  // Helper method to validate token and refresh
  // if needed
  function getAccessToken(callback) {
    var now = new Date().getTime();
    var isExpired = now > parseInt(sessionStorage.tokenExpires);
    // Do we have a token already?
    if (sessionStorage.accessToken && !isExpired) {
      // Just return what we have
      if (callback) {
        callback(sessionStorage.accessToken);
      }
    } else {
      // Attempt to do a hidden iframe request
      makeSilentTokenRequest(callback);
    }
  }

  // OUTLOOK API FUNCTIONS =======================

  function makeApiCall(options, callback) {
    var headers = {
      // Add the access token to the request
      'Authorization': 'Bearer ' + options.token,
      // Set a request ID (helpful for troubleshooting)
      'client-request-id': guid(),
      // Request that the client request ID be returned
      'return-client-request-id': `true`
    };

    // If specified, set the user's email as the anchor
    // This helps API requests route to the appropriate server
    // more efficiently
    if (options.email) {
      headers['X-AnchorMailbox'] = options.email;
    }

    var ajaxOptions = {
      url: options.url,
      dataType: 'json',
      type: options.method,
      headers: headers
    };

    if (options.query) {
      ajaxOptions['data'] = options.query;
    }

    $.ajax(ajaxOptions)
    .done(function(response) {
      callback(response);
    })
    .fail(function(error) {
      callback(null, error);
    });
  }

  function getUserEmailAddress(callback) {
    if (sessionStorage.userEmail) {
      return sessionStorage.userEmail;
    } else {
      getAccessToken(function(accessToken) {
        if (accessToken) {
          // Call the Outlook API /Me to get user email address
          var callOptions = {
            url: apiEndpoint + '/Me',
            token: accessToken,
            method: 'GET'
          }; 

          makeApiCall(callOptions, function(result, error) {
            if (error) {
              callback(null, error);
            } else {
              callback(result.EmailAddress);
            }
          });
        } else {
          var error = { responseText: 'Could not retrieve access token' };
          callback(null, error);
        }
      });
    }
  }

  function getUserInboxMessages(emailAddress, callback) {
    getAccessToken(function(accessToken) {
      if (accessToken) {
        // Call the Outlook API
        var callOptions = {
          // Get messages from the inbox folder
          url: apiEndpoint + '/Me/mailfolders/inbox/messages',
          token: accessToken,
          method: 'GET',
          email: emailAddress,
          query: {
            // Limit to the first 10 messages
            '$top': 10,
            // Only return fields we will use
            '$select': 'Subject,From,ReceivedDateTime,BodyPreview',
            // Sort by received time, newest first
            '$orderby': 'ReceivedDateTime DESC'
          }
        }; 

        makeApiCall(callOptions, function(result, error) {
          if (error) {
            callback(null, error);
          } else {
            callback(result.value);
          }
        });
      } else {
        var error = { responseText: 'Could not retrieve access token' };
        callback(null, error);
      }
    });
  }

  function getUserEvents(emailAddress, callback) {
    getAccessToken(function(accessToken) {
      if (accessToken) {
        // Call the Outlook API
        var callOptions = {
          // Get events from the calendar
          url: apiEndpoint + '/Me/events',
          token: accessToken,
          method: 'GET',
          email: emailAddress,
          query: {
            // Limit to the first 10 events
            '$top': 10,
            // Only return fields we will use
            '$select': 'Subject,Start,End,CreatedDateTime',
            // Sort by created time
            '$orderby': 'CreatedDateTime DESC'
          }
        }; 

        makeApiCall(callOptions, function(result, error) {
          if (error) {
            callback(null, error);
          } else {
            callback(result.value);
          }
        });
      } else {
        var error = { responseText: 'Could not retrieve access token' };
        callback(null, error);
      }
    });
  }

  function getUserContacts(emailAddress, callback) {
    getAccessToken(function(accessToken) {
      if (accessToken) {
        // Call the Outlook API
        var callOptions = {
          // Get contacts
          url: apiEndpoint + '/Me/contacts',
          token: accessToken,
          method: 'GET',
          email: emailAddress,
          query: {
            // Limit to the first 10 contacts
            '$top': 10,
            // Only return fields we will use
            '$select': 'GivenName,Surname,EmailAddresses',
            // Sort by given name alphabetically
            '$orderby': 'GivenName ASC'
          }
        }; 

        makeApiCall(callOptions, function(result, error) {
          if (error) {
            callback(null, error);
          } else {
            callback(result.value);
          }
        });
      } else {
        var error = { responseText: 'Could not retrieve access token' };
        callback(null, error);
      }
    });
  }

  // HELPER FUNCTIONS ============================

  function guid() {
    function s4() {
      return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
    }
    return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
      s4() + '-' + s4() + s4() + s4();
  }

  function parseHashParams(hash) {
    var params = hash.slice(1).split('&');
    
    var paramarray = {};
    params.forEach(function(param) {
      param = param.split('=');
      paramarray[param[0]] = param[1];
    });
    
    return paramarray;
  }

  function decodePlusEscaped(value) {
    // decodeURIComponent doesn't handle spaces escaped
    // as '+'
    if (value) {
      return decodeURIComponent(value.replace(/\+/g, ' '));
    } else {
      return '';
    }
  }

  function clearUserState() {
    // Clear session
    sessionStorage.clear();
  }

  Handlebars.registerHelper("formatDate", function(datetime){
    // Dates from API look like:
    // 2016-06-27T14:06:13Z

    var date = new Date(datetime);
    return date.toLocaleDateString() + ' ' + date.toLocaleTimeString();
  });
});