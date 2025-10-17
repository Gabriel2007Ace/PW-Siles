// Este script lida com a lógica da página de exportação de planilhas.
document.addEventListener('DOMContentLoaded', function() {
    const exportPage = document.getElementById('exportacao-planilha');
    if (exportPage) {
        const filterForm = document.getElementById('filter-form');
        const pedidosExportTableBody = document.getElementById('pedidos-export-table-body');
        const noPedidosExportMessage = document.getElementById('no-pedidos-export-message');
        const selectAllCheckbox = document.getElementById('select-all-pedidos');
        const generateSelectedPlanilhaBtn = document.getElementById('generate-selected-planilha-btn');
        const exportMessage = document.getElementById('export-message');

        // --- Funções Auxiliares ---

        // Função para mostrar mensagens na página de exportação
        function showExportMessage(message, isSuccess = true) {
            exportMessage.innerText = message;
            exportMessage.classList.remove('hidden');
            if (isSuccess) {
                exportMessage.classList.add('text-green-600');
                exportMessage.classList.remove('text-red-600');
            } else {
                exportMessage.classList.add('text-red-600');
                exportMessage.classList.remove('text-green-600');
            }
        }

        // Função para carregar e exibir os pedidos na tabela de exportação
        async function loadPedidosForExport(filters = {}) {
            pedidosExportTableBody.innerHTML = ''; // Limpa a tabela
            noPedidosExportMessage.classList.add('hidden'); // Esconde a mensagem
            selectAllCheckbox.checked = false; // Desmarca o "selecionar todos"

            const userId = localStorage.getItem('userId');
            if (!userId) {
                showExportMessage('Você precisa estar logado para ver os pedidos.', false);
                return;
            }

            try {
                // CORREÇÃO AQUI: Filtra apenas pedidos com status 'confirmado' para produção
                // Removido 'producao' se não for usado
                const defaultFilters = { status: 'confirmado' }; // Altere para 'confirmado' se 'producao' não existe
                
                // Se o filtro de status for 'Todos', não o inclua nos filtros
                if (filters.status === 'Todos os Status' || filters.status === '') {
                    delete filters.status;
                }

                const mergedFilters = { ...defaultFilters, ...filters };
                // Se um status específico foi selecionado no filtro, use-o
                if (filters.status && filters.status !== 'Todos os Status') {
                    mergedFilters.status = filters.status;
                } else if (filters.status === 'Todos os Status') {
                    delete mergedFilters.status; // Remove o filtro de status se "Todos" for selecionado
                }


                const queryParams = new URLSearchParams(mergedFilters).toString();
                const response = await fetch(`/api/pedidos?${queryParams}`, { // Usa a API de listagem de pedidos
                    method: 'GET',
                    headers: { 'X-User-Id': userId }
                });

                const pedidos = await response.json();

                if (pedidos.length > 0) {
                    pedidos.forEach(pedido => {
                        const row = document.createElement('tr');
                        row.className = 'table-row';
                        
                        let statusClass = '';
                        if (pedido.status === 'confirmado') statusClass = 'status-confirmado';
                        else if (pedido.status === 'pendente') statusClass = 'status-pendente';
                        else if (pedido.status === 'producao') statusClass = 'status-producao'; // Manter se você adicionar este status

                        row.innerHTML = `
                            <td class="px-6 py-4 whitespace-nowrap">
                                <input type="checkbox" class="pedido-checkbox form-checkbox h-4 w-4 text-blue-600 transition duration-150 ease-in-out" data-pedido-id="${pedido.id}">
                            </td>
                            <td class="px-6 py-4 whitespace-nowrap">${pedido.clienteNome}</td>
                            <td class="px-6 py-4 whitespace-nowrap">${pedido.dataRetirada}</td>
                            <td class="px-6 py-4 whitespace-nowrap">${pedido.tipoPedido}</td>
                            <td class="px-6 py-4 whitespace-nowrap">${pedido.quantidade}</td>
                            <td class="px-6 py-4 whitespace-nowrap">
                                <span class="${statusClass}">${pedido.status}</span>
                            </td>
                        `;
                        pedidosExportTableBody.appendChild(row);
                    });
                } else {
                    noPedidosExportMessage.innerText = 'Nenhum pedido encontrado para os filtros.';
                    noPedidosExportMessage.classList.remove('hidden');
                }

            } catch (error) {
                console.error('Erro ao carregar pedidos para exportação:', error);
                showExportMessage('Erro ao carregar pedidos. Tente novamente.', false);
                noPedidosExportMessage.classList.remove('hidden');
            }
        }

        // --- Event Listeners ---

        // Carrega pedidos ao carregar a página (sem filtros iniciais)
        window.addEventListener('load', () => {
            if (window.location.pathname === '/exportar') {
                loadPedidosForExport(); // Carrega todos os pedidos inicialmente
            }
        });

        // Lógica para aplicar filtros
        filterForm.addEventListener('submit', function(event) {
            event.preventDefault();
            const formData = new FormData(filterForm);
            const filters = Object.fromEntries(formData.entries());
            loadPedidosForExport(filters);
        });

        // Lógica do checkbox "Selecionar Todos"
        selectAllCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.pedido-checkbox');
            checkboxes.forEach(checkbox => {
                checkbox.checked = selectAllCheckbox.checked;
            });
        });

        // Lógica para o botão "Gerar Planilha dos Selecionados"
        generateSelectedPlanilhaBtn.addEventListener('click', async function() {
            const selectedPedidoIds = [];
            document.querySelectorAll('.pedido-checkbox:checked').forEach(checkbox => {
                selectedPedidoIds.push(checkbox.dataset.pedidoId);
            });

            if (selectedPedidoIds.length === 0) {
                showExportMessage('Por favor, selecione pelo menos um pedido para exportar.', false);
                return;
            }

            const userId = localStorage.getItem('userId');
            if (!userId) {
                showModal('Você precisa estar logado para exportar uma planilha.'); // showModal do main.js
                return;
            }

            try {
                const response = await fetch('/api/reports/export-selected-pedidos', { // Nova rota POST
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-User-Id': userId
                    },
                    body: JSON.stringify({ pedido_ids: selectedPedidoIds })
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = `planilha_producao_selecionados_${new Date().toISOString().slice(0,10)}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    
                    showExportMessage('Planilha gerada com sucesso! O download deve ter sido iniciado.', true);
                } else {
                    const errorData = await response.json();
                    showExportMessage(errorData.message || 'Erro ao gerar a planilha dos selecionados.', false);
                }
            } catch (error) {
                console.error('Erro de conexão ao gerar planilha:', error);
                showExportMessage('Erro de conexão com o servidor. Tente novamente.', false);
            }
        });
    }
});