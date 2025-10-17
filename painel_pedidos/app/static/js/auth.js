document.addEventListener('DOMContentLoaded', function() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const targetTab = this.dataset.tab;
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(targetTab).classList.add('active');
        });
    });


    // =========================================================================
    // ========= INÍCIO DA NOVA LÓGICA DE COMUNICAÇÃO COM O BACKEND =========
    // =========================================================================

    // --- LÓGICA DO FORMULÁRIO DE LOGIN ---
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const form = e.target;
            const formData = new FormData(form);
            const data = Object.fromEntries(formData.entries());
            const loginError = document.getElementById('login-error');

            try {
                const response = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });

                const result = await response.json();

                if (response.ok) {
                    loginError.style.display = 'none';
                    // SALVA O ID DO USUÁRIO NO NAVEGADOR E REDIRECIONA
                    localStorage.setItem('userId', result.userId);
                    window.location.href = '/pedidos'; // Redireciona para o painel
                } else {
                    loginError.textContent = result.message || 'Erro no login.';
                    loginError.style.display = 'block';
                }
            } catch (error) {
                loginError.textContent = 'Erro de conexão com o servidor.';
                loginError.style.display = 'block';
            }
        });
    }

    // --- LÓGICA DO FORMULÁRIO DE CADASTRO ---
    const registerForm = document.getElementById('registerForm');
    if (registerForm) {
        registerForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const password = document.getElementById('password').value;
            const confirmPassword = document.getElementById('confirmPassword').value;
            
            if (password !== confirmPassword) {
                alert('As senhas não coincidem!');
                return;
            }

            const form = e.target;
            const formData = new FormData(form);
            const data = Object.fromEntries(formData.entries());
            
            const registerError = document.getElementById('register-error');
            const registerSuccess = document.getElementById('register-success');

            try {
                const response = await fetch('/api/auth/register', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });

                const result = await response.json();

                if (response.ok) {
                    registerError.style.display = 'none';
                    registerSuccess.textContent = result.message + ' Por favor, faça login para continuar.';
                    registerSuccess.style.display = 'block';
                    form.reset(); // Limpa o formulário
                } else {
                    registerSuccess.style.display = 'none';
                    registerError.textContent = result.message || 'Erro ao cadastrar.';
                    registerError.style.display = 'block';
                }
            } catch (error) {
                registerSuccess.style.display = 'none';
                registerError.textContent = 'Erro de conexão com o servidor.';
                registerError.style.display = 'block';
            }
        });
    }
    



// INICIO DA LOGICA "INUTIL POR ENQUANTO"



 const passwordToggles = document.querySelectorAll('.password-toggle');
    passwordToggles.forEach(toggle => {
        toggle.addEventListener('click', function() {
            const input = this.previousElementSibling;
            const icon = this.querySelector('i');

            if (input.type === 'password') {
                input.type = 'text';
                icon.classList.remove('fa-eye');
                icon.classList.add('fa-eye-slash');
            } else {
                input.type = 'password';
                icon.classList.remove('fa-eye-slash');
                icon.classList.add('fa-eye');
            }
        });
    });

    const passwordInput = document.getElementById('password');
    if (passwordInput) {
        passwordInput.addEventListener('input', function() {
            const password = this.value;
            const strengthFill = document.querySelector('.strength-fill');
            const strengthText = document.querySelector('.strength-text');

            if (password.length === 0) {
                strengthFill.className = 'strength-fill';
                strengthText.textContent = '';
                return;
            }

            let strength = 0;
            if (password.length >= 8) strength++;
            if (password.match(/[a-z]/) && password.match(/[A-Z]/)) strength++;
            if (password.match(/[0-9]/)) strength++;
            if (password.match(/[^a-zA-Z0-9]/)) strength++;

            strengthFill.className = 'strength-fill';
            if (strength <= 1) {
                strengthFill.classList.add('weak');
                strengthText.textContent = 'Senha fraca';
                strengthText.style.color = 'var(--danger-color)';
            } else if (strength <= 3) {
                strengthFill.classList.add('medium');
                strengthText.textContent = 'Senha média';
                strengthText.style.color = 'var(--warning-color)';
            } else {
                strengthFill.classList.add('strong');
                strengthText.textContent = 'Senha forte';
                strengthText.style.color = 'var(--success-color)';
            }
        });
    }

    const confirmPasswordInput = document.getElementById('confirmPassword');
    if (confirmPasswordInput) {
        confirmPasswordInput.addEventListener('input', function() {
            const password = document.getElementById('password').value;
            const confirmPassword = this.value;
            const feedback = this.parentElement.nextElementSibling;

            if (confirmPassword === '') {
                feedback.textContent = '';
                feedback.className = 'input-feedback';
                return;
            }

            if (password === confirmPassword) {
                feedback.textContent = 'As senhas coincidem';
                feedback.className = 'input-feedback success';
            } else {
                feedback.textContent = 'As senhas não coincidem';
                feedback.className = 'input-feedback error';
            }
        });
    }

    const emailInput = document.querySelector('.email-input');
    if (emailInput) {
        emailInput.addEventListener('blur', function() {
            const email = this.value;
            const feedback = this.nextElementSibling;
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

            if (email === '') {
                feedback.textContent = '';
                feedback.className = 'input-feedback';
                return;
            }

            if (emailRegex.test(email)) {
                feedback.textContent = 'E-mail válido';
                feedback.className = 'input-feedback success';
            } else {
                feedback.textContent = 'E-mail inválido';
                feedback.className = 'input-feedback error';
            }
        });
    }})