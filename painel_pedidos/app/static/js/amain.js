// Este código JavaScript foi adaptado para a arquitetura de Multi-Page Application (MPA)
// As funções de showScreen() que alternavam entre divs foram removidas,
// pois a navegação agora é feita pelo Flask e pelas URLs.

// --- Funções de Modal (Pop-ups) ---
// Elas continuam úteis para exibir mensagens de feedback ao usuário.
const messageModal = document.getElementById('messageModal');
const modalCloseButton = document.getElementById('modal-close-button');
const modalOkButton = document.getElementById('modal-ok-button');

function showModal(message) {
    document.getElementById('modalMessage').innerText = message;
    messageModal.style.display = 'flex';
}

function closeModal() {
    messageModal.style.display = 'none';
}

if (modalCloseButton) {
    modalCloseButton.addEventListener('click', closeModal);
}
if (modalOkButton) {
    modalOkButton.addEventListener('click', closeModal);
}


// --- Lógica de Autenticação e Navegação ---

// Lidar com o formulário de REGISTRO
const registerForm = document.getElementById('register-form');
if (registerForm) { // O código só roda se o formulário existir na página atual (ex: /login)
    registerForm.addEventListener('submit', async function(event) {
        event.preventDefault();

        const email = document.getElementById('register-email').value;
        const password = document.getElementById('register-password').value;
        const registerError = document.getElementById('register-error');

        try {
            const response = await fetch('/api/auth/register', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email, password })
            });

            const data = await response.json();

            if (response.ok) {
                console.log('Cadastro bem-sucedido!', data.message);
                registerError.classList.add('hidden');
                showModal(data.message + ' Agora, faça login.');
                window.location.href = '/login'; // Redireciona para a página de login
            } else {
                console.error('Erro no cadastro:', data.message);
                registerError.innerText = data.message || 'Erro ao cadastrar usuário.';
                registerError.classList.remove('hidden');
            }
        } catch (error) {
            console.error('Erro de conexão ou inesperado durante o registro:', error);
            registerError.innerText = 'Erro ao conectar com o servidor. Tente novamente.';
            registerError.classList.remove('hidden');
        }
    });
}

// Lidar com o formulário de LOGIN
const loginForm = document.getElementById('login-form');
if (loginForm) { // O código só roda se o formulário existir na página atual (ex: /login)
    loginForm.addEventListener('submit', async function(event) {
        event.preventDefault();

        const email = document.getElementById('login-email').value;
        const password = document.getElementById('login-password').value;
        const loginError = document.getElementById('login-error');

        try {
            const response = await fetch('/api/auth/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email, password })
            });

            const data = await response.json();

            if (response.ok) {
                console.log('Login bem-sucedido!', data.message);
                loginError.classList.add('hidden');
                localStorage.setItem('userId', data.userId);
                // Redireciona para a página principal de pedidos após o login
                window.location.href = '/pedidos';
            } else {
                console.error('Erro no login:', data.message);
                loginError.innerText = data.message || 'E-mail ou senha incorretos.';
                loginError.classList.remove('hidden');
            }
        } catch (error) {
            console.error('Erro de conexão ou inesperado durante o login:', error);
            loginError.innerText = 'Erro ao conectar com o servidor. Tente novamente.';
            loginError.classList.remove('hidden');
        }
    });

    // Lidar com o botão de alternância entre Login e Cadastro
    document.getElementById('show-register').addEventListener('click', function() {
        document.getElementById('login').classList.remove('active');
        document.getElementById('cadastro-usuario').classList.add('active');
    });

    document.getElementById('show-login').addEventListener('click', function() {
        document.getElementById('cadastro-usuario').classList.remove('active');
        document.getElementById('login').classList.add('active');
    });
}

// Lidar com o botão de LOGOUT
const logoutButton = document.getElementById('logout-button');
if (logoutButton) {
    logoutButton.addEventListener('click', function() {
        localStorage.removeItem('userId'); // Remove o ID do usuário do armazenamento local
        showModal('Você foi desconectado.');
        window.location.href = '/login'; // Redireciona para a página de login
    });
}

// Lógica de verificação de autenticação ao carregar qualquer página
window.addEventListener('load', () => {
    const userId = localStorage.getItem('userId');
    const navbar = document.getElementById('navbar');
    if (navbar) {
        if (userId) {
            navbar.classList.remove('hidden'); // Mostra a navbar se o usuário estiver logado
        } else {
            navbar.classList.add('hidden'); // Esconde a navbar se o usuário não estiver logado
            // Se a página não for de login, redireciona para o login
            if (window.location.pathname !== '/login' && window.location.pathname !== '/') {
                window.location.href = '/login';
            }
        }
    }
});




// --- Lógica para o formulário de Cadastro de Pedido ---
// --- Lógica para o formulário de Cadastro de Pedido ---
// --- Lógica para o formulário de Cadastro de Pedido ---
const pedidoForm = document.getElementById('pedido-form');
if (pedidoForm) {
    pedidoForm.addEventListener('submit', async function(event) {
        event.preventDefault(); // Evita o envio padrão do formulário

        const userId = localStorage.getItem('userId'); // Pega o userId do armazenamento local

        if (!userId) {
            // Se não houver userId, o usuário não está logado
            const pedidoError = document.getElementById('pedido-error');
            pedidoError.innerText = 'Você precisa estar logado para criar um pedido.';
            pedidoError.classList.remove('hidden');
            return;
        }

        const form = event.target;
        const formData = new FormData(form);
        const pedidoData = Object.fromEntries(formData.entries());

        const pedidoError = document.getElementById('pedido-error');
        const pedidoSuccess = document.getElementById('pedido-success');

        try {
            const response = await fetch('/api/pedidos', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-User-Id': userId // Adiciona o ID do usuário ao cabeçalho da requisição
                },
                body: JSON.stringify(pedidoData)
            });

            const data = await response.json();

            if (response.ok) {
                console.log('Pedido salvo com sucesso!', data);
                pedidoError.classList.add('hidden');
                pedidoSuccess.innerText = 'Pedido salvo com sucesso!';
                pedidoSuccess.classList.remove('hidden');

                form.reset();

                setTimeout(() => {
                    window.location.href = '/pedidos'; // Redireciona para a lista de pedidos
                }, 2000);

            } else {
                console.error('Erro ao salvar pedido:', data.message);
                pedidoSuccess.classList.add('hidden');
                pedidoError.innerText = data.message || 'Erro ao salvar pedido. Verifique os dados.';
                pedidoError.classList.remove('hidden');
            }
        } catch (error) {
            console.error('Erro de conexão ou inesperado:', error);
            pedidoSuccess.classList.add('hidden');
            pedidoError.innerText = 'Erro ao conectar com o servidor. Tente novamente.';
            pedidoError.classList.remove('hidden');
        }
    });
}


// --- Funções para a Listagem de Pedidos ---

// Função para mostrar pedidos
// Função para buscar e exibir os pedidos
// --- Funções para a Listagem de Pedidos ---

// Função para buscar e exibir os pedidos
async function loadPedidos(filters = {}) {
    const tableBody = document.getElementById('pedidos-table-body');
    const noPedidosMessage = document.getElementById('no-pedidos-message');
    const userId = localStorage.getItem('userId');

    // Limpa a tabela e esconde a mensagem de "nenhum pedido" antes de carregar
    tableBody.innerHTML = '';
    noPedidosMessage.classList.add('hidden');

    if (!userId) {
        // Se o usuário não estiver logado, não busca nada e exibe uma mensagem
        noPedidosMessage.innerText = 'Faça login para ver os pedidos.';
        noPedidosMessage.classList.remove('hidden');
        return;
    }

    // --- INÍCIO DA MUDANÇA ---
    // Constrói a URL com os parâmetros de filtro
    const params = new URLSearchParams();
    if (filters.dataInicio) params.append('dataInicio', filters.dataInicio);
    if (filters.dataFim) params.append('dataFim', filters.dataFim);
    if (filters.tipoProduto) params.append('tipoProduto', filters.tipoProduto);

    const queryString = params.toString();
    const apiUrl = `/api/pedidos?${queryString}`; // A URL agora pode ter filtros
    // --- FIM DA MUDANÇA ---

    try {
        const response = await fetch('/api/pedidos', {
            method: 'GET',
            headers: {
                'X-User-Id': userId // Envia o ID do usuário para o backend
            }
        });

        const pedidos = await response.json();

        if (pedidos.length > 0) {
            // Se houver pedidos, itera sobre eles e cria uma linha na tabela para cada um
            pedidos.forEach(pedido => {
                const row = document.createElement('tr');
                row.className = 'table-row';
                
                // Define a classe do status com base no valor retornado pelo backend
                let statusClass = '';
                if (pedido.status === 'confirmado') {
                    statusClass = 'status-confirmado';
                } else if (pedido.status === 'pendente') {
                    statusClass = 'status-pendente';
                } else if (pedido.status === 'producao') {
                    statusClass = 'status-producao';
                }

                let prioridadeClass = pedido.prioridade === 'urgente' ? 'text-red-600 font-bold' : '';


                row.innerHTML = `
                    <td class="px-6 py-4 whitespace-nowrap">${pedido.clienteNome}</td>
                    <td class="px-6 py-4 whitespace-nowrap">${pedido.dataEvento}</td>
                    <td class="px-6 py-4 whitespace-nowrap">${pedido.tipoPedido}</td>
                    <td class="px-6 py-4 whitespace-nowrap">${pedido.quantidade}</td>
                    <td class="px-6 py-4 whitespace-nowrap ${prioridadeClass}">${pedido.prioridade || 'normal'}</td>

                    <td class="px-6 py-4 whitespace-nowrap">
                        <span class="${statusClass}">${pedido.status}</span>
                    </td>
                    <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                        <a href="/pedidos/${pedido.id}" class="text-blue-600 hover:text-blue-900 mr-2">Ver Detalhes</a>
                        <button class="confirm-pedido-btn text-green-600 hover:text-green-900 mr-2" data-pedido-id="${pedido.id}">Confirmar</button>
                        <button class="delete-pedido-btn text-red-600 hover:text-red-900" data-pedido-id="${pedido.id}">Excluir</button>
                    </td>
                `;
                tableBody.appendChild(row);
            });
        } else {
            // Se não houver pedidos, exibe a mensagem de "nenhum pedido"
            noPedidosMessage.innerText = 'Nenhum pedido encontrado. Crie um novo pedido.';
            noPedidosMessage.classList.remove('hidden');
        }

    } catch (error) {
        console.error('Erro ao carregar pedidos:', error);
        noPedidosMessage.innerText = 'Erro ao carregar pedidos. Tente novamente mais tarde.';
        noPedidosMessage.classList.remove('hidden');
    }
}

// Chama a função para carregar os pedidos quando a página for carregada
window.addEventListener('load', () => {
    // ... (restante da lógica de verificação de autenticação no load) ...
    
    // NOVO: Chama a função loadPedidos() se a página atual for a de listagem
    if (window.location.pathname === '/pedidos') {
        loadPedidos();
    }
});



// logica da tela de detalhes de pedido

// --- Lógica para a Página de Detalhes do Pedido ---

// --- Lógica para a Página de Detalhes do Pedido ---

async function loadPedidoDetails(pedidoId) {
    const userId = localStorage.getItem('userId');
    if (!userId) {
        console.error("Usuário não logado. Redirecionando para login.");
        window.location.href = '/login';
        return;
    }

    try {
        const response = await fetch(`/api/pedidos/${pedidoId}`, {
            method: 'GET',
            headers: { 'X-User-Id': userId }
        });

        if (response.ok) {
            const pedido = await response.json();
            console.log("Dados do pedido carregados:", pedido);

            // Preenche a página com os dados do pedido (campos de visualização)
            document.getElementById('detalhe-pedido-id').innerText = `#${pedidoId}`;
            document.getElementById('detalhe-cliente-nome').innerText = pedido.clienteNome || 'N/A';
            document.getElementById('detalhe-cliente-telefone').innerText = pedido.clienteTelefone || 'N/A';
            document.getElementById('detalhe-cliente-email').innerText = pedido.clienteEmail || 'N/A';
            document.getElementById('detalhe-observacoes-gerais-cliente').innerText = pedido.observacoesGeraisCliente || 'N/A';
            document.getElementById('detalhe-tipo-pedido').innerText = pedido.tipoPedido || 'N/A';
            document.getElementById('detalhe-data-evento').innerText = pedido.dataEvento || 'N/A';
            document.getElementById('detalhe-data-retirada').innerText = pedido.dataRetirada || 'N/A';
            document.getElementById('detalhe-horario-retirada').innerText = pedido.horarioRetirada || 'N/A';
            document.getElementById('detalhe-quantidade').innerText = pedido.quantidade || 'N/A';
            document.getElementById('detalhe-tipo-embalagem').innerText = pedido.tipoEmbalagem || 'N/A';
            document.getElementById('detalhe-sabores').innerText = pedido.sabores || 'N/A';
            document.getElementById('detalhe-observacoes').innerText = pedido.observacoes || 'N/A';

            // NOVOS CAMPOS DO CONTRATO (visualização)
            document.getElementById('detalhe-cliente-rg').innerText = pedido.clienteRG || 'N/A';
            document.getElementById('detalhe-cliente-cpf').innerText = pedido.clienteCPF || 'N/A';
            document.getElementById('detalhe-nome-contratado').innerText = pedido.nomeContratado || 'N/A';
            document.getElementById('detalhe-cnpj-contratado').innerText = pedido.cnpjContratado || 'N/A';
            document.getElementById('detalhe-valor-total-contrato').innerText = pedido.valorTotalPedidoContrato || 'N/A';
            document.getElementById('detalhe-data-pagamento-contrato').innerText = pedido.dataPagamentoContrato || 'N/A';
            document.getElementById('detalhe-local-evento').innerText = pedido.localEvento || 'N/A';


            // Atualiza o status visual
            const statusSpan = document.getElementById('detalhe-status-text');
            const statusIndicator = document.getElementById('detalhe-status-indicator');
            statusSpan.innerText = pedido.status;
            statusIndicator.className = statusIndicator.className.split(' ').filter(c => !c.startsWith('status-')).join(' '); // Limpa status anterior
            if (pedido.status === 'confirmado') statusIndicator.classList.add('status-confirmado');
            else if (pedido.status === 'pendente') statusIndicator.classList.add('status-pendente');
            else if (pedido.status === 'producao') statusIndicator.classList.add('status-producao');

            // --- Preencher Tabela de Produtos Contratados ---
            const produtosTableBody = document.getElementById('produtos-contratados-table-body');
            const noProdutosContratadosMessage = document.getElementById('no-produtos-contratados-message');
            produtosTableBody.innerHTML = ''; // Limpa qualquer conteúdo anterior
            noProdutosContratadosMessage.classList.add('hidden'); // Esconde a mensagem por padrão

            try {
                const produtosContratados = JSON.parse(pedido.produtosContratadosJson || '[]');
                if (produtosContratados.length > 0) {
                    produtosContratados.forEach(produto => {
                        const row = document.createElement('tr');
                        row.innerHTML = `
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${produto.Quantidade || 'N/A'}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${produto.Produto || 'N/A'}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${produto['Valor Unitário'] || 'N/A'}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${produto['Valor Total Item'] || 'N/A'}</td>
                        `;
                        produtosTableBody.appendChild(row);
                    });
                } else {
                    noProdutosContratadosMessage.classList.remove('hidden');
                }
            } catch (jsonError) {
                console.error("Erro ao analisar JSON de produtos contratados:", jsonError);
                noProdutosContratadosMessage.innerText = "Erro ao carregar produtos contratados.";
                noProdutosContratadosMessage.classList.remove('hidden');
            }

        } else {
            console.error("Erro ao carregar detalhes do pedido:", await response.json());
            showModal('Tela de Cadastro de Pedidos!');
        }
    } catch (error) {
        console.error("Erro de conexão ao buscar pedido:", error);
        showModal('Erro de conexão com o servidor ao carregar detalhes do pedido.');
    }
}
// Chama a função para carregar os detalhes do pedido quando a página de detalhes for carregada
window.addEventListener('load', () => {
    const path = window.location.pathname;
    const detailsPageMatch = path.match(/\/pedidos\/(.*)/);
    if (detailsPageMatch) {
        const pedidoId = detailsPageMatch[1];
        loadPedidoDetails(pedidoId);
    }
});



// --- Lógica para os botões de Ação na Lista de Pedidos ---

async function handlePedidoAction(action, pedidoId) {
    const userId = localStorage.getItem('userId');
    if (!userId) {
        console.error("Usuário não logado. Redirecionando para login.");
        window.location.href = '/login';
        return;
    }

    let method = '';
    let body = null;
    let endpoint = `/api/pedidos/${pedidoId}`;

    if (action === 'confirm') {
        method = 'PUT';
        body = JSON.stringify({ status: 'confirmado' });
    } else if (action === 'delete') {
        method = 'DELETE';
    } else {
        console.error("Ação desconhecida:", action);
        return;
    }

    try {
        const response = await fetch(endpoint, {
            method: method,
            headers: {
                'Content-Type': 'application/json',
                'X-User-Id': userId
            },
            body: body
        });

        if (response.ok) {
            const result = await response.json();
            console.log(`Pedido ${action === 'confirm' ? 'confirmado' : 'excluído'} com sucesso!`, result.message);
            // Após a ação, recarrega a lista de pedidos para refletir a mudança
            loadPedidos();
        } else {
            const error = await response.json();
            console.error(`Erro ao ${action} pedido:`, error.message);
            showModal(`Erro ao ${action} pedido: ${error.message}`);
        }
    } catch (error) {
        console.error("Erro de conexão ao processar a ação:", error);
        showModal("Erro de conexão com o servidor. Tente novamente.");
    }
}

// Adiciona os event listeners aos botões de ação
document.addEventListener('click', function(event) {
    const target = event.target;
    const pedidoId = target.dataset.pedidoId;

    if (target.classList.contains('confirm-pedido-btn') && pedidoId) {
        if (confirm("Tem certeza que deseja confirmar este pedido?")) {
            handlePedidoAction('confirm', pedidoId);
        }
    } else if (target.classList.contains('delete-pedido-btn') && pedidoId) {
        if (confirm("Tem certeza que deseja excluir este pedido? Esta ação não pode ser desfeita.")) {
            handlePedidoAction('delete', pedidoId);
        }
    }
});

// --- logica do upload de contrato começa aqui ---//

// --- Lógica para a página de Upload de Contratos ---
const uploadForm = document.getElementById('upload-contratos');
if (uploadForm) {
    const fileInput = document.getElementById('contrato-upload');
    const fileNameSpan = document.getElementById('file-name');
    const selectFileBtn = document.getElementById('select-file-btn');
    const analisarBtn = document.getElementById('analisar-contrato-btn');
    const previewDiv = document.getElementById('extraido-preview');
    const extractedDataDisplay = document.getElementById('extracted-data-display');
    const uploadError = document.getElementById('upload-error');
    const uploadSuccess = document.getElementById('upload-success');

    // Abre a janela de seleção de arquivo ao clicar no botão
    selectFileBtn.addEventListener('click', () => {
        fileInput.click();
    });

    // Exibe o nome do arquivo selecionado
    fileInput.addEventListener('change', () => {
        const file = fileInput.files[0];
        if (file) {
            fileNameSpan.innerText = file.name;
        } else {
            fileNameSpan.innerText = '';
        }
    });

    // Envia o arquivo para o backend e exibe os dados extraídos
    analisarBtn.addEventListener('click', async function() {
        const file = fileInput.files[0];
        if (!file) {
            uploadError.innerText = 'Por favor, selecione um arquivo para análise.';
            uploadError.classList.remove('hidden');
            return;
        }

        const userId = localStorage.getItem('userId');
        if (!userId) {
            showModal('Você precisa estar logado para analisar um contrato.');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        uploadError.classList.add('hidden');
        uploadSuccess.classList.add('hidden');
        previewDiv.classList.add('hidden');

        try {
            const response = await fetch('/api/contracts/upload', {
                method: 'POST',
                headers: {
                    'X-User-Id': userId,
                },
                body: formData
            });

            const result = await response.json();
            if (response.ok) {
                console.log('Contrato processado e pedido salvo!', result.pedido);
            showModal(result.message); // Mostra o pop-up com "Contrato processado e pedido salvo com sucesso!"

            // Espera 2 segundos e redireciona para a lista de pedidos
            setTimeout(() => {
                window.location.href = '/pedidos'; 
            }, 2000);
                // Salva os dados extraídos em uma variável global para uso posterior
                window.extractedData = result.extractedData;

            } else {
                uploadError.innerText = result.message || 'Erro ao analisar o contrato.';
                uploadError.classList.remove('hidden');
            }
        } catch (error) {
            console.error('Erro ao analisar contrato:', error);
            uploadError.innerText = 'Erro de conexão com o servidor. Tente novamente.';
            uploadError.classList.remove('hidden');
        }
    });

//     // Lógica para salvar os dados extraídos como um novo pedido
//     const saveAsPedidoForm = document.getElementById('save-pedido-from-contract-form');
//     if (saveAsPedidoForm) {
//         saveAsPedidoForm.addEventListener('submit', async function(event) {
//             event.preventDefault();

//             const userId = localStorage.getItem('userId');
//             if (!userId) {
//                 showModal('Você precisa estar logado para salvar um pedido.');
//                 return;
//             }

//             // Pega os dados que foram salvos do passo anterior
//             const pedidoData = window.extractedData;
//             if (!pedidoData) {
//                 showModal('Nenhum dado de contrato foi extraído para ser salvo.');
//                 return;
//             }
            
//             // Adiciona campos que não vieram do extrator mas são necessários
//             pedidoData.status = 'pendente'; 
            
//             // Reutiliza a lógica de criação de pedido
//             try {
//                 const response = await fetch('/api/pedidos', {
//                     method: 'POST',
//                     headers: {
//                         'Content-Type': 'application/json',
//                         'X-User-Id': userId,
//                     },
//                     body: JSON.stringify(pedidoData)
//                 });
//                 const result = await response.json();
//                 if (response.ok) {
//                     showModal('Pedido salvo com sucesso a partir do contrato! Redirecionando...');
//                     setTimeout(() => {
//                         window.location.href = '/pedidos';
//                     }, 2000);
//                 } else {
//                     showModal(`Erro ao salvar pedido do contrato: ${result.message}`);
//                 }
//             } catch (error) {
//                 showModal('Erro de conexão ao salvar pedido.');
//                 console.error('Erro ao salvar pedido do contrato:', error);
//             }
//         });
//     }
// }

// --- Lógica para salvar os dados extraídos como um novo pedido ---
const saveAsPedidoForm = document.getElementById('save-pedido-from-contract-form');
if (saveAsPedidoForm) {
    saveAsPedidoForm.addEventListener('submit', async function(event) {
        event.preventDefault();

        const userId = localStorage.getItem('userId');
        if (!userId) {
            showModal('Você precisa estar logado para salvar um pedido.');
            return;
        }

        const pedidoData = window.extractedData;
        if (!pedidoData) {
            showModal('Nenhum dado de contrato foi extraído para ser salvo.');
            return;
        }

        // CORREÇÃO AQUI: Preencher campos obrigatórios que podem estar faltando
        // O Extractor.py não preenche dataRetirada e horarioRetirada
        // Vamos usar a data do evento como valor padrão para a data de retirada por enquanto
        // ou você pode adicionar um formulário na página para o usuário preencher.
        pedidoData.dataRetirada = pedidoData.dataRetirada || pedidoData.dataEvento;
        pedidoData.horarioRetirada = pedidoData.horarioRetirada || '12:00'; // Valor padrão
        
        // Adiciona o status, que é obrigatório no backend
        pedidoData.status = 'pendente';
        
        try {
            const response = await fetch('/api/pedidos', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-User-Id': userId,
                },
                body: JSON.stringify(pedidoData)
            });
            const result = await response.json();
            if (response.ok) {
                showModal('Pedido salvo com sucesso a partir do contrato! Redirecionando...');
                setTimeout(() => {
                    window.location.href = '/pedidos';
                }, 2000);
            } else {
                showModal(`Erro ao salvar pedido do contrato: ${result.message}`);
            }
        } catch (error) {
            showModal('Erro de conexão ao salvar pedido.');
            console.error('Erro ao salvar pedido do contrato:', error);
        }
    });
}}





// --- Lógica para a Edição de Pedidos ---

const detalhesPedidoPage = document.getElementById('detalhes-pedido');
if (detalhesPedidoPage) {
    const btnEditarPedido = document.getElementById('btn-editar-pedido');
    const btnSalvarEdicao = document.getElementById('btn-salvar-edicao');
    const btnCancelarEdicao = document.getElementById('btn-cancelar-edicao');
    const adminActions = document.getElementById('admin-actions');
    
    // Funções auxiliares para alternar o modo de visualização/edição
    function toggleEditMode(isEditing) {
        // Alterna a visibilidade dos campos de visualização (p) e edição (input/textarea/select)
        document.querySelectorAll('[id^="detalhe-"]').forEach(el => {
            el.classList.toggle('hidden', isEditing);
        });
        document.querySelectorAll('.editable-field').forEach(el => {
            el.classList.toggle('hidden', !isEditing);
        });

        if (isEditing) {
            // Entra no modo de edição: Mostra os botões de salvar/cancelar
            adminActions.classList.remove('hidden');
            btnEditarPedido.classList.add('hidden');
        } else {
            // Sai do modo de edição: Esconde os botões de salvar/cancelar
            adminActions.classList.add('hidden');
            btnEditarPedido.classList.remove('hidden');
        }
    }
    
    // Preenche os campos de input com os dados do pedido atual para edição
    // Preenche os campos de input com os dados do pedido atual para edição
function fillEditFields(pedido) {
    // Campos gerais do pedido
    document.getElementById('edit-cliente-nome').value = pedido.clienteNome || '';
    document.getElementById('edit-cliente-telefone').value = pedido.clienteTelefone || '';
    document.getElementById('edit-cliente-email').value = pedido.clienteEmail || '';
    document.getElementById('edit-observacoes-gerais-cliente').value = pedido.observacoesGeraisCliente || '';
    document.getElementById('edit-tipo-pedido').value = pedido.tipoPedido || '';
    document.getElementById('edit-data-evento').value = pedido.dataEvento || '';
    document.getElementById('edit-data-retirada').value = pedido.dataRetirada || '';
    document.getElementById('edit-horario-retirada').value = pedido.horarioRetirada || '';
    document.getElementById('edit-quantidade').value = pedido.quantidade || '';
    document.getElementById('edit-tipo-embalagem').value = pedido.tipoEmbalagem || '';
    document.getElementById('edit-sabores').value = pedido.sabores || '';
    document.getElementById('edit-observacoes').value = pedido.observacoes || '';

    // NOVOS CAMPOS DO CONTRATO para edição
    document.getElementById('edit-cliente-rg').value = pedido.clienteRG || '';
    document.getElementById('edit-cliente-cpf').value = pedido.clienteCPF || '';
    document.getElementById('edit-nome-contratado').value = pedido.nomeContratado || '';
    document.getElementById('edit-cnpj-contratado').value = pedido.cnpjContratado || '';
    document.getElementById('edit-valor-total-contrato').value = pedido.valorTotalPedidoContrato || '';
    document.getElementById('edit-data-pagamento-contrato').value = pedido.dataPagamentoContrato || '';
    document.getElementById('edit-local-evento').value = pedido.localEvento || '';
}

    // Ação do botão "Editar Pedido"
    btnEditarPedido.addEventListener('click', function() {
        // Encontra o ID do pedido na URL
        const pedidoId = window.location.pathname.split('/').pop();
        
        const userId = localStorage.getItem('userId');
        if (!userId) {
            showModal('Você precisa estar logado para editar um pedido.');
            return;
        }

        // Faz uma requisição para buscar os dados do pedido
        fetch(`/api/pedidos/${pedidoId}`, {
            method: 'GET',
            headers: { 'X-User-Id': userId }
        }).then(response => response.json())
        .then(pedido => {
            fillEditFields(pedido); // Preenche os campos de edição
            toggleEditMode(true); // Entra no modo de edição
        }).catch(error => {
            console.error('Erro ao buscar dados do pedido para edição:', error);
            showModal('Não foi possível carregar os dados para edição.');
        });
    });

    // Ação do botão "Salvar Alterações"
    btnSalvarEdicao.addEventListener('click', async function() {
        const pedidoId = window.location.pathname.split('/').pop();
        const userId = localStorage.getItem('userId');
        
        const updatedData = {
            clienteNome: document.getElementById('edit-cliente-nome').value,
            clienteTelefone: document.getElementById('edit-cliente-telefone').value,
            clienteEmail: document.getElementById('edit-cliente-email').value,
            observacoesGeraisCliente: document.getElementById('edit-observacoes-gerais-cliente').value,
            tipoPedido: document.getElementById('edit-tipo-pedido').value,
            dataEvento: document.getElementById('edit-data-evento').value,
            dataRetirada: document.getElementById('edit-data-retirada').value,
            horarioRetirada: document.getElementById('edit-horario-retirada').value,
            quantidade: document.getElementById('edit-quantidade').value,
            tipoEmbalagem: document.getElementById('edit-tipo-embalagem').value,
            sabores: document.getElementById('edit-sabores').value,
            observacoes: document.getElementById('edit-observacoes').value
        };

        try {
            const response = await fetch(`/api/pedidos/${pedidoId}`, {
                method: 'PUT', // Requisição PUT para atualizar
                headers: {
                    'Content-Type': 'application/json',
                    'X-User-Id': userId
                },
                body: JSON.stringify(updatedData)
            });
            const result = await response.json();
            if (response.ok) {
                showModal('Pedido atualizado com sucesso!');
                toggleEditMode(false); // Volta para o modo de visualização
                // Opcional: recarregar a página para mostrar os novos dados
                window.location.reload(); 
            } else {
                showModal('Erro ao salvar alterações: ' + result.message);
            }
        } catch (error) {
            console.error('Erro de conexão ao salvar alterações:', error);
            showModal('Erro de conexão com o servidor. Tente novamente.');
        }
    });

    // Ação do botão "Cancelar Edição"
    btnCancelarEdicao.addEventListener('click', function() {
        toggleEditMode(false); // Volta para o modo de visualização
    });
}



// --- Lógica para o botão de Gerar Comprovante de Retirada ---
const btnGerarComprovante = document.getElementById('btn-gerar-comprovante');
if (btnGerarComprovante) {
    btnGerarComprovante.addEventListener('click', async function() {
        const pedidoId = window.location.pathname.split('/').pop();
        const userId = localStorage.getItem('userId');
        
        if (!userId) {
            showModal('Você precisa estar logado para gerar um comprovante.');
            return;
        }

        try {
            const response = await fetch(`/api/reports/generate-delivery-report/${pedidoId}`, {
                method: 'GET',
                headers: {
                    'X-User-Id': userId
                }
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `comprovante_retirada_${pedidoId}.docx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                showModal('Comprovante gerado com sucesso! O download deve ter sido iniciado.');
            } else {
                const error = await response.json();
                showModal(`Erro ao gerar comprovante: ${error.message}`);
                console.error("Erro ao gerar comprovante:", error);
            }
        } catch (error) {
            console.error("Erro de conexão ao gerar comprovante:", error);
            showModal('Erro de conexão com o servidor. Tente novamente.');
        }
    });
}


// Nova logica de cadastro do pedido
// --- Lógica para a Seção Opcional de Informações do Contrato (Novo Pedido) ---
const toggleContractInfo = document.getElementById('toggle-contract-info');
if (toggleContractInfo) {
    const contractInfoFields = document.getElementById('contract-info-fields');
    const toggleIcon = document.getElementById('toggle-icon');

    toggleContractInfo.addEventListener('click', function() {
        contractInfoFields.classList.toggle('hidden'); // Alterna a visibilidade
        if (contractInfoFields.classList.contains('hidden')) {
            toggleIcon.innerText = '▼'; // Seta para baixo quando escondido
        } else {
            toggleIcon.innerText = '▲'; // Seta para cima quando visível
        }
    });
}


// --- Lógica para a Página de Relatórios ---
// --- Lógica de Paineis de Relatorios
async function loadRelatorios() {
    const userId = localStorage.getItem('userId');
    if (!userId) {
        console.error("Usuário não logado. Redirecionando para login.");
        window.location.href = '/login';
        return;
    }

    try {
        const response = await fetch('/api/reports', { // Rota da API de relatórios
            method: 'GET',
            headers: {
                'X-User-Id': userId
            }
        });

        if (response.ok) {
            const relatoriosData = await response.json();
            console.log("Dados de relatórios carregados:", relatoriosData);

            // Preenche os cards de indicadores
            document.getElementById('total-pedidos-mes').innerText = relatoriosData.totalPedidosMes || '0';
            document.getElementById('produtos-mais-pedidos').innerText = relatoriosData.produtosMaisPedidos || 'N/A';
            document.getElementById('clientes-mais-pedidos').innerText = relatoriosData.clientesMaisPedidos || 'N/A';

            // Preenche o gráfico de evolução (agora com Chart.js)
            const evolucaoChartCanvas = document.getElementById('evolucao-chart');
            if (evolucaoChartCanvas) { // Verifica se o canvas existe na página
                // Verifica se já existe uma instância do gráfico e a destrói para evitar conflitos
                if (window.myEvolutionChart instanceof Chart) {
                    window.myEvolutionChart.destroy();
                }

                // Rótulos para os períodos (ex: meses ou semanas)
                // Você pode ajustar estes rótulos para refletir seus dados reais, se quiser.
                const labels = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']; 
                const dataValues = relatoriosData.evolucaoSemanalMensal || []; // Pega os valores do backend

                // Cria uma nova instância do gráfico
                window.myEvolutionChart = new Chart(evolucaoChartCanvas, {
                    type: 'line', // Tipo de gráfico: linha
                    data: {
                        labels: labels.slice(0, dataValues.length), // Usa apenas os rótulos correspondentes aos dados
                        datasets: [{
                            label: 'Evolução de Pedidos',
                            data: dataValues,
                            borderColor: 'rgb(75, 192, 192)',
                            backgroundColor: 'rgba(75, 192, 192, 0.2)', // Adiciona um preenchimento sutil
                            tension: 0.4, // Suaviza a linha do gráfico
                            fill: true, // Preenche a área abaixo da linha
                            pointBackgroundColor: 'rgb(75, 192, 192)',
                            pointBorderColor: '#fff',
                            pointHoverBackgroundColor: '#fff',
                            pointHoverBorderColor: 'rgb(75, 192, 192)'
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false, // Permite que o gráfico se ajuste ao tamanho do contêiner
                        plugins: {
                            legend: {
                                display: true,
                                position: 'top',
                            },
                            title: {
                                display: true,
                                text: 'Evolução Semanal/Mensal de Pedidos'
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: true, // Começa o eixo Y em zero
                                title: {
                                    display: true,
                                    text: 'Quantidade de Pedidos'
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: 'Período'
                                }
                            }
                        }
                    }
                });
            }


        } else {
            const error = await response.json();
            console.error("Erro ao carregar relatórios:", error.message);
            showModal(`Erro ao carregar relatórios: ${error.message}`);
        }
    } catch (error) {
        console.error("Erro de conexão ao buscar relatórios:", error);
        showModal('Erro de conexão com o servidor ao carregar relatórios.');
    }
}

// Chama a função para carregar os relatórios quando a página for carregada
window.addEventListener('load', () => {
    // ... (restante da lógica de verificação de autenticação no load) ...
    
    // NOVO: Chama a função loadRelatorios() se a página atual for a de relatórios
    if (window.location.pathname === '/relatorios') {
        loadRelatorios();
    }
});


// --- Lógica para o Formulário de Criação de Contrato ---

const contratoForm = document.getElementById('contrato-form');
if (contratoForm) {
    const produtosTableBody = document.getElementById('produtos-contratados-table-body');
    const addProductBtn = document.getElementById('add-product-btn');

    // Função para adicionar uma nova linha à tabela de produtos
    function addProductRow() {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                <input type="number" name="quantidade" required class="input-field text-center">
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                <input type="text" name="produto" required class="input-field">
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                <input type="text" name="valorUnitario" required class="input-field text-center">
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                <input type="text" name="valorTotalItem" required class="input-field text-center">
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                <button type="button" class="remove-product-btn text-red-600 hover:text-red-900">Remover</button>
            </td>
        `;
        produtosTableBody.appendChild(row);
    }
    
    // Adiciona a primeira linha por padrão
    addProductRow();

    // Adiciona um event listener para o botão "Adicionar Produto"
    addProductBtn.addEventListener('click', addProductRow);

    // Lógica para remover uma linha da tabela
    produtosTableBody.addEventListener('click', function(event) {
        if (event.target.classList.contains('remove-product-btn')) {
            const row = event.target.closest('tr');
            row.remove();
        }
    });

    // Lógica para enviar o formulário e gerar o contrato
    contratoForm.addEventListener('submit', async function(event) {
        event.preventDefault();
        
        const userId = localStorage.getItem('userId');
        if (!userId) {
            showModal('Você precisa estar logado para gerar um contrato.');
            return;
        }

        const formData = new FormData(contratoForm);
        const contratoData = Object.fromEntries(formData.entries());

        // Coleta os dados da tabela dinâmica de produtos
        const produtos = [];
        document.querySelectorAll('#produtos-contratados-table-body tr').forEach(row => {
            const inputs = row.querySelectorAll('input');
            const produto = {
                'Quantidade': inputs[0].value,
                'Produto': inputs[1].value,
                'Valor Unitário': inputs[2].value,
                'Valor Total Item': inputs[3].value
            };
            produtos.push(produto);
        });
        contratoData.produtosContratados = produtos;

       // console.log("Enviando para o backend:", JSON.stringify(contratoData, null, 2));


        try {
            const response = await fetch('/api/contracts/gerar-contrato', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-User-Id': userId,
                },
                body: JSON.stringify(contratoData)
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `contrato_gerado_${new Date().toISOString().slice(0, 10)}.docx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                showModal('Contrato gerado com sucesso! O download deve ter sido iniciado.');
                contratoForm.reset(); // Limpa o formulário após o sucesso
            } else {
                const error = await response.json();
                showModal(`Erro ao gerar contrato: ${error.message}`);
                console.error("Erro ao gerar contrato:", error);
            }
        } catch (error) {
            console.error('Erro de conexão ao gerar contrato:', error);
            showModal('Erro de conexão com o servidor. Tente novamente.');
        }
    });
}



// --- Lógica para a Calculadora de Custos ---

const calculadoraPage = document.getElementById('calculadora-custos');
if (calculadoraPage) {
    // Dados de exemplo simulando um banco de dados
    const recipes = [
        {
            id: 'brigadeiro-beijinho',
            name: 'Brigadeiro de Beijinho',
            sellingPrice: 40.00,
            yields: 30, // 30 unidades
            totalWeight: 500, // 500g
            prepTime: 45, // 45 minutos
            ingredients: [
                { name: 'Leite Condensado', quantity: 395, unit: 'g', defaultPricePerUnit: 20.00 },
                { name: 'Manteiga', quantity: 20, unit: 'g', defaultPricePerUnit: 30.00 },
                { name: 'Coco Ralado', quantity: 100, unit: 'g', defaultPricePerUnit: 40.00 },
                { name: 'Açúcar Cristal', quantity: 30, unit: 'g', defaultPricePerUnit: 5.00 },
                { name: 'Cravo-da-Índia (opc.)', quantity: 10, unit: 'g', defaultPricePerUnit: 50.00 }
            ]
        },
        {
            id: 'trufa-morango',
            name: 'Trufa de Morango',
            sellingPrice: 55.00,
            yields: 25,
            totalWeight: 450,
            prepTime: 60,
            ingredients: [
                { name: 'Chocolate Meio Amargo', quantity: 200, unit: 'g', defaultPricePerUnit: 60.00 },
                { name: 'Creme de Leite', quantity: 100, unit: 'ml', defaultPricePerUnit: 15.00 },
                { name: 'Morango Fresco', quantity: 150, unit: 'g', defaultPricePerUnit: 25.00 },
                { name: 'Confeitos de Chocolate', quantity: 50, unit: 'g', defaultPricePerUnit: 35.00 }
            ]
        },
        {
            id: 'bolo-chocolate',
            name: 'Bolo de Chocolate',
            sellingPrice: 80.00,
            yields: 16,
            totalWeight: 1200,
            prepTime: 90,
            ingredients: [
                { name: 'Farinha de Trigo', quantity: 300, unit: 'g', defaultPricePerUnit: 8.00 },
                { name: 'Açúcar', quantity: 250, unit: 'g', defaultPricePerUnit: 6.00 },
                { name: 'Chocolate em Pó', quantity: 100, unit: 'g', defaultPricePerUnit: 25.00 },
                { name: 'Ovos', quantity: 3, unit: 'un', defaultPricePerUnit: 0.80 },
                { name: 'Leite', quantity: 200, unit: 'ml', defaultPricePerUnit: 5.00 },
                { name: 'Óleo', quantity: 100, unit: 'ml', defaultPricePerUnit: 12.00 }
            ]
        }
    ];

    // Referências aos elementos do DOM
    const recipeSelect = document.getElementById('receita-select');
    const insumosTableBody = document.getElementById('insumos-table-body');
    const insumosSection = document.getElementById('insumos-section');
    const resultadosSection = document.getElementById('resultados-section');
    
    // Elementos de resultados
    const rendimentoUnidades = document.getElementById('rendimento-unidades');
    const pesoTotalGramas = document.getElementById('peso-total-gramas');
    const tempoPreparo = document.getElementById('tempo-preparo');
    const custoTotalReceita = document.getElementById('custo-total-receita');
    const custoPorUnidade = document.getElementById('custo-por-unidade');
    const precoModelo = document.getElementById('preco-modelo');
    const margemLucro = document.getElementById('margem-lucro');
    const margemIndicator = document.getElementById('margem-indicator');
    
    // --- Funções da Calculadora ---

    // Função para carregar as opções do dropdown de receitas
    function loadRecipesIntoDropdown() {
        recipeSelect.innerHTML = '<option value="">Selecione uma receita...</option>';
        recipes.forEach(recipe => {
            const option = document.createElement('option');
            option.value = recipe.id;
            option.textContent = recipe.name;
            recipeSelect.appendChild(option);
        });
    }

    // Função para calcular o custo de um ingrediente
    function calculateIngredientCost(pricePerUnit, quantity, unit) {
        let cost = 0;
        if (unit === 'g' || unit === 'ml') {
            cost = (pricePerUnit * quantity) / 1000;
        } else if (unit === 'un') {
            cost = pricePerUnit * quantity;
        }
        return cost;
    }

    // Função para carregar os ingredientes da receita selecionada na tabela
    function loadIngredientsIntoTable(selectedRecipeId) {
        insumosTableBody.innerHTML = '';
        const selectedRecipe = recipes.find(r => r.id === selectedRecipeId);

        if (!selectedRecipe) {
            insumosSection.classList.add('hidden');
            resultadosSection.classList.add('hidden');
            return;
        }

        rendimentoUnidades.textContent = `${selectedRecipe.yields} unidades`;
        pesoTotalGramas.textContent = `${selectedRecipe.totalWeight} g`;
        tempoPreparo.textContent = `${selectedRecipe.prepTime} min`;

        selectedRecipe.ingredients.forEach(ingredient => {
            const row = document.createElement('tr');
            row.className = 'hover:bg-gray-50';
            
            const initialCost = calculateIngredientCost(ingredient.defaultPricePerUnit, ingredient.quantity, ingredient.unit);

            row.innerHTML = `
                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${ingredient.name}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${ingredient.quantity}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${ingredient.unit}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                    <input type="number"
                            class="ingredient-price-input w-28 p-2 border border-gray-300 rounded-md shadow-sm focus:border-blue-500 focus:ring focus:ring-blue-500 focus:ring-opacity-50 transition-all duration-200 text-gray-800"
                            value="${ingredient.defaultPricePerUnit.toFixed(2)}"
                            step="0.01"
                            data-quantity="${ingredient.quantity}"
                            data-unit="${ingredient.unit}">
                </td>
                <td class="ingredient-calculated-cost px-6 py-4 whitespace-nowrap text-sm text-gray-900 font-semibold">R$ ${initialCost.toFixed(2)}</td>
            `;
            insumosTableBody.appendChild(row);
        });

        addPriceInputListeners();
        calculateAllCosts(selectedRecipe);

        insumosSection.classList.remove('hidden');
        resultadosSection.classList.remove('hidden');
    }

    // Função para adicionar listeners aos inputs de preço dos ingredientes
    function addPriceInputListeners() {
        document.querySelectorAll('.ingredient-price-input').forEach(input => {
            input.addEventListener('input', (event) => {
                const pricePerUnit = parseFloat(event.target.value) || 0;
                const quantity = parseFloat(event.target.dataset.quantity);
                const unit = event.target.dataset.unit;
                const calculatedCostElement = event.target.closest('tr').querySelector('.ingredient-calculated-cost');

                const newCost = calculateIngredientCost(pricePerUnit, quantity, unit);
                calculatedCostElement.textContent = `R$ ${newCost.toFixed(2)}`;

                const selectedRecipeId = recipeSelect.value;
                const selectedRecipe = recipes.find(r => r.id === selectedRecipeId);
                calculateAllCosts(selectedRecipe);
            });
        });
    }

    // Função para calcular o custo total, custo por unidade e a margem de lucro
    function calculateAllCosts(selectedRecipe) {
        let totalProductionCost = 0;
        document.querySelectorAll('.ingredient-calculated-cost').forEach(costElement => {
            const costText = costElement.textContent;
            const costValue = parseFloat(costText.replace('R$', '').replace(',', '.'));
            if (!isNaN(costValue)) {
                totalProductionCost += costValue;
            }
        });

        const custoPorUnidadeCalc = totalProductionCost / selectedRecipe.yields;

        custoTotalReceita.textContent = `R$ ${totalProductionCost.toFixed(2).replace('.', ',')}`;
        custoPorUnidade.textContent = `R$ ${custoPorUnidadeCalc.toFixed(2).replace('.', ',')}`;
        precoModelo.textContent = `R$ ${selectedRecipe.sellingPrice.toFixed(2).replace('.', ',')}`;
        
        let profitMargin = 0;
        if (selectedRecipe.sellingPrice > 0) {
            profitMargin = ((selectedRecipe.sellingPrice - custoPorUnidadeCalc) / selectedRecipe.sellingPrice) * 100;
        }

        margemLucro.textContent = `${profitMargin.toFixed(2)} %`;

        if (profitMargin >= 40) {
            margemIndicator.style.backgroundColor = '#10B981';
        } else {
            margemIndicator.style.backgroundColor = '#EF4444';
        }
    }
    
    // --- Lógica para o modal de Adicionar Receita ---
    const addRecipeModal = document.getElementById('add-recipe-modal');
    const openAddRecipeBtn = document.getElementById('open-add-recipe-modal-btn');
    const closeRecipeModalBtn = document.getElementById('close-recipe-modal');
    const newRecipeForm = document.getElementById('new-recipe-form');
    const addIngredientBtn = document.getElementById('add-new-ingredient-btn');
    const newRecipeTableBody = document.getElementById('new-recipe-ingredients-table');

    openAddRecipeBtn.addEventListener('click', () => {
        addRecipeModal.classList.remove('hidden');
    });

    closeRecipeModalBtn.addEventListener('click', () => {
        addRecipeModal.classList.add('hidden');
    });
    
    // Adiciona linha de ingrediente no modal
    addIngredientBtn.addEventListener('click', () => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900"><input type="text" name="name" required class="input-field"></td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900"><input type="number" name="quantity" required class="input-field w-16"></td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900"><input type="text" name="unit" required class="input-field w-16"></td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900"><input type="number" name="defaultPricePerUnit" step="0.01" required class="input-field w-24"></td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-center"><button type="button" class="remove-ingredient-btn text-red-600 hover:text-red-900">Remover</button></td>
        `;
        newRecipeTableBody.appendChild(row);
    });

    // Remove linha de ingrediente no modal
    newRecipeTableBody.addEventListener('click', (event) => {
        if (event.target.classList.contains('remove-ingredient-btn')) {
            event.target.closest('tr').remove();
        }
    });

    // Lógica para salvar a nova receita (Apenas em memória para o MVP)
    newRecipeForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const formData = new FormData(newRecipeForm);
        const newRecipe = {};
        const ingredients = [];
        let currentIngredient = {};
        
        for (const [key, value] of formData.entries()) {
            if (['name', 'quantity', 'unit', 'defaultPricePerUnit'].includes(key)) {
                currentIngredient[key] = value;
                if (Object.keys(currentIngredient).length === 4) {
                    ingredients.push(currentIngredient);
                    currentIngredient = {};
                }
            } else {
                newRecipe[key] = value;
            }
        }
        
        newRecipe.id = 'new-recipe-' + Date.now();
        newRecipe.ingredients = ingredients;
        recipes.push(newRecipe);

        showModal('Nova receita cadastrada com sucesso!');
        newRecipeForm.reset();
        addRecipeModal.classList.add('hidden');
        loadRecipesIntoDropdown();
        recipeSelect.value = newRecipe.id;
        loadIngredientsIntoTable(newRecipe.id);
    });

    // --- Inicialização ---
    loadRecipesIntoDropdown();
    const firstRecipe = recipes[0];
    if (firstRecipe) {
        recipeSelect.value = firstRecipe.id;
        loadIngredientsIntoTable(firstRecipe.id);
    }
}



// logica de filtros para a pagina de exportar:

// Em static/js/main.js (pode adicionar no final)

const aplicarFiltrosBtn = document.getElementById('apply-export-filters-btn');

if (aplicarFiltrosBtn) {
    aplicarFiltrosBtn.addEventListener('click', function() {
        // Pega os valores dos campos de filtro
        const dataInicio = document.getElementById('filter-data-inicio').value;
        const dataFim = document.getElementById('filter-data-fim').value;
        const tipoProduto = document.getElementById('filter-tipo-produto').value;

        // Monta o objeto de filtros
        const filters = {
            dataInicio: dataInicio,
            dataFim: dataFim,
            tipoProduto: tipoProduto
        };

        // Chama a função loadPedidos com os filtros!
        loadPedidos(filters);
    });
}