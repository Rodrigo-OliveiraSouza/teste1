function normalizeWhitespace(value) {
  return value.replace(/\s+/g, ' ').trim();
}

function formatStatus(kind, message) {
  const prefix = {
    info: 'Info',
    success: 'Ok',
    error: 'Erro'
  }[kind] || 'Info';
  return `${prefix}: ${message}`;
}

function assertElement(element, name) {
  if (!element) {
    throw new Error(`Missing ${name}`);
  }
  return element;
}

function parseJwtPayload(token) {
  try {
    const payload = token.split('.')[1];
    if (!payload) {
      return null;
    }
    const base64 = payload.replace(/-/g, '+').replace(/_/g, '/');
    const json = atob(base64);
    const decoded = decodeURIComponent(
      Array.from(json)
        .map((char) => `%${char.charCodeAt(0).toString(16).padStart(2, '0')}`)
        .join('')
    );
    return JSON.parse(decoded);
  } catch (error) {
    return null;
  }
}

export function initUI() {
  const form = assertElement(document.getElementById('contact-form'), '#contact-form');
  const status = assertElement(document.getElementById('contact-status'), '#contact-status');
  const submitButton = assertElement(form.querySelector('button[type="submit"]'), 'contact submit button');

  const nameInput = assertElement(document.getElementById('contact-name'), '#contact-name');
  const messageInput = assertElement(document.getElementById('contact-message'), '#contact-message');
  const honeypotInput = assertElement(document.getElementById('contact-company'), '#contact-company');
  const googleButton = assertElement(document.getElementById('google-signin'), '#google-signin');
  const googleStatus = assertElement(document.getElementById('google-status'), '#google-status');
  const googleSignout = assertElement(document.getElementById('google-signout'), '#google-signout');

  const state = {
    csrfToken: '',
    csrfExpiresAt: 0,
    sending: false
  };

  const auth = {
    idToken: '',
    email: '',
    name: ''
  };

  const defaultAuthText = 'Faca login com Google para enviar sua ideia.';

  const setStatus = (kind, message) => {
    status.textContent = formatStatus(kind, message);
  };

  const updateSubmitState = () => {
    submitButton.disabled = state.sending || !auth.idToken;
  };

  const setAuth = ({ idToken = '', email = '', name = '' } = {}) => {
    auth.idToken = idToken;
    auth.email = email;
    auth.name = name;

    if (auth.idToken) {
      googleStatus.textContent = `Conectado como ${auth.email || 'sua conta Google'}.`;
      googleSignout.classList.add('is-visible');
      if (!nameInput.value && auth.name) {
        nameInput.value = auth.name;
      }
    } else {
      googleStatus.textContent = defaultAuthText;
      googleSignout.classList.remove('is-visible');
    }

    updateSubmitState();
  };

  const setAuthUnavailable = (message) => {
    googleStatus.textContent = message;
    googleSignout.classList.remove('is-visible');
    updateSubmitState();
  };

  const waitForGoogle = () => new Promise((resolve) => {
    if (window.google?.accounts?.id) {
      resolve(window.google);
      return;
    }
    const startedAt = Date.now();
    const interval = setInterval(() => {
      if (window.google?.accounts?.id) {
        clearInterval(interval);
        resolve(window.google);
        return;
      }
      if (Date.now() - startedAt > 5000) {
        clearInterval(interval);
        resolve(null);
      }
    }, 200);
  });

  const initGoogle = async () => {
    const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID;
    if (!clientId) {
      setAuthUnavailable('Login Google indisponivel no momento.');
      return;
    }

    const googleApi = await waitForGoogle();
    if (!googleApi) {
      setAuthUnavailable('Login Google indisponivel no momento.');
      return;
    }

    const handleCredential = (response) => {
      if (!response?.credential) {
        setAuth();
        return;
      }
      const payload = parseJwtPayload(response.credential) || {};
      setAuth({
        idToken: response.credential,
        email: payload.email || '',
        name: payload.name || payload.given_name || ''
      });
    };

    googleApi.accounts.id.initialize({
      client_id: clientId,
      callback: handleCredential
    });

    googleApi.accounts.id.renderButton(googleButton, {
      theme: 'outline',
      size: 'large',
      text: 'continue_with',
      shape: 'pill'
    });
  };

  googleSignout.addEventListener('click', () => {
    if (window.google?.accounts?.id) {
      if (auth.email) {
        window.google.accounts.id.revoke(auth.email, () => {
          setAuth();
        });
      } else {
        window.google.accounts.id.disableAutoSelect();
        setAuth();
      }
    } else {
      setAuth();
    }
  });

  const fetchCsrf = async () => {
    const now = Date.now();
    if (state.csrfToken && state.csrfExpiresAt - 5000 > now) {
      return state.csrfToken;
    }

    const response = await fetch('/api/csrf', {
      method: 'GET',
      headers: {
        Accept: 'application/json'
      },
      credentials: 'same-origin'
    });

    if (!response.ok) {
      throw new Error('csrf_failed');
    }

    const data = await response.json();
    state.csrfToken = data.token;
    state.csrfExpiresAt = data.expiresAt || 0;
    return state.csrfToken;
  };

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (state.sending) {
      return;
    }

    if (!auth.idToken) {
      setStatus('error', 'Faca login com Google para enviar a mensagem.');
      return;
    }

    const nameValue = normalizeWhitespace(nameInput.value) || auth.name;
    const payload = {
      name: nameValue,
      message: messageInput.value.trim(),
      company: honeypotInput.value.trim(),
      idToken: auth.idToken
    };

    if (!payload.name || !payload.message) {
      setStatus('error', 'Preencha nome e mensagem.');
      return;
    }

    state.sending = true;
    updateSubmitState();
    setStatus('info', 'Enviando...');

    try {
      const csrfToken = await fetchCsrf();
      const response = await fetch('/api/contact', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-CSRF-Token': csrfToken
        },
        body: JSON.stringify(payload),
        credentials: 'same-origin'
      });

      if (response.ok) {
        setStatus('success', 'Mensagem enviada com sucesso.');
        form.reset();
        setAuth({ idToken: auth.idToken, email: auth.email, name: auth.name });
      } else if (response.status === 429) {
        setStatus('error', 'Limite de envios atingido. Tente mais tarde.');
      } else if (response.status === 403) {
        setStatus('error', 'Sessao expirada. Recarregue a pagina.');
      } else if (response.status === 401) {
        setStatus('error', 'Faca login com Google novamente.');
      } else {
        setStatus('error', 'Nao foi possivel enviar agora.');
      }
    } catch (error) {
      setStatus('error', 'Falha de rede ao enviar.');
    } finally {
      state.sending = false;
      updateSubmitState();
    }
  });

  initGoogle();
  updateSubmitState();

  return {
    setStatus
  };
}
