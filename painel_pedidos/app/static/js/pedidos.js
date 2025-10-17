const userId = localStorage.getItem("userId");

document.addEventListener("DOMContentLoaded", () => {
    carregarPedidos();

    const formFiltros = document.getElementById("filtrosPedidos");
    formFiltros.addEventListener("submit", (e) => {
        e.preventDefault();
        carregarPedidos();
    });
});

async function carregarPedidos() {
    const tbody = document.getElementById("pedidosTableBody");
    const mensagemErro = document.getElementById("mensagemErro");
    tbody.innerHTML = "";

    // --- Coleta dos filtros
    const cliente = document.getElementById("clienteFiltro").value.trim();
    const dataEvento = document.getElementById("filtroData").value; // YYYY-MM-DD from input
    const status = document.getElementById("filtroStatus").value;

    // Função utilitária: converte "YYYY-MM-DD" -> "DD/MM/YYYY"
    const yyyyToDdMmYyyy = (isoDate) => {
        if (!isoDate) return "";
        const parts = isoDate.split("-");
        if (parts.length !== 3) return isoDate;
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    };

    // --- Monta URL com parâmetros
    const params = new URLSearchParams();
    if (cliente) params.append("cliente", cliente);
    if (dataEvento) params.append("dataEvento", dataEvento);
    if (status) params.append("status", status);

    try {
        console.log("Buscando pedidos com params:", params.toString());
        const response = await fetch(`/api/pedidos?${params.toString()}`, {
            headers: {
                "Content-Type": "application/json",
                "X-User-Id": userId
            }
        });

        if (!response.ok) throw new Error(`Erro na resposta da API: ${response.status}`);
        let pedidos = await response.json();

        // Se veio vazio e havia filtro de data, tenta o fallback com DD/MM/YYYY
        if ((Array.isArray(pedidos) && pedidos.length === 0) && dataEvento) {
            const altParams = new URLSearchParams();
            if (cliente) altParams.append("cliente", cliente);
            altParams.append("dataEvento", yyyyToDdMmYyyy(dataEvento));
            if (status) altParams.append("status", status);

            console.log("Nenhum resultado. Tentando fallback com formato DD/MM/YYYY:", altParams.toString());
            const altResp = await fetch(`/api/pedidos?${altParams.toString()}`, {
                headers: {
                    "Content-Type": "application/json",
                    "X-User-Id": userId
                }
            });

            if (altResp.ok) {
                const altPedidos = await altResp.json();
                if (Array.isArray(altPedidos) && altPedidos.length > 0) {
                    console.log("Fallback retornou resultados.");
                    pedidos = altPedidos;
                } else {
                    console.log("Fallback não retornou resultados.");
                }
            } else {
                console.warn("Fallback retornou erro:", altResp.status);
            }
        }

        if (!Array.isArray(pedidos) || pedidos.length === 0) {
            tbody.innerHTML = `<tr><td colspan="9" class="text-center">Nenhum pedido encontrado.</td></tr>`;
            return;
        }

        pedidos.forEach(pedido => {
            const statusCapitalizado = pedido.status ? pedido.status.charAt(0).toUpperCase() + pedido.status.slice(1) : "-";
            const prioridadeCapitalizada = pedido.prioridade ? pedido.prioridade.charAt(0).toUpperCase() + pedido.prioridade.slice(1) : "Normal";
            const responsavel = pedido.responsavel || "Damaris";

            const row = document.createElement("tr");
            row.classList.add("align-middle");

            row.innerHTML = `
                <td>${pedido.id}</td>
                <td>${pedido.clienteNome || "-"}</td>
                <td>${pedido.dataEvento || "-"}</td>
                <td>${pedido.tipoPedido || "-"}</td>
                <td>${pedido.quantidade ?? "-"}</td>
                <td><span class="badge ${pedido.status === 'confirmado' ? 'badge-success' : (pedido.status === 'producao' ? 'badge-warning' : 'badge-warning')}">${statusCapitalizado}</span></td>
                <td><span class="badge ${pedido.prioridade === 'alta' || pedido.prioridade === 'urgente' ? 'badge-danger' : 'badge-secondary'}">${prioridadeCapitalizada}</span></td>
                <td>${responsavel}</td>
                <td class="text-center">
                    <button class="btn btn-sm btn-outline-primary btn-detalhes me-2" title="Ver detalhes">
                        <i class="fas fa-eye"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-success btn-confirmar me-2" title="Confirmar pedido">
                        <i class="fas fa-check"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-danger btn-excluir" title="Excluir pedido">
                        <i class="fas fa-trash-alt"></i>
                    </button>
                </td>
                `;


            // Listeners confirmar/excluir
            row.querySelector(".btn-confirmar").addEventListener("click", async () => {
                if (!confirm("Deseja confirmar este pedido?")) return;
                try {
                    const res = await fetch(`/api/pedidos/${pedido.id}`, {
                        method: "PUT",
                        headers: {
                            "Content-Type": "application/json",
                            "X-User-Id": userId
                        },
                        body: JSON.stringify({ status: "confirmado" })
                    });
                    if (!res.ok) throw new Error();
                    alert("Pedido confirmado!");
                    carregarPedidos();
                } catch {
                    alert("Erro ao confirmar pedido.");
                }
            });

            row.querySelector(".btn-excluir").addEventListener("click", async () => {
                if (!confirm("Tem certeza que deseja excluir este pedido?")) return;
                try {
                    const res = await fetch(`/api/pedidos/${pedido.id}`, {
                        method: "DELETE",
                        headers: {
                            "Content-Type": "application/json",
                            "X-User-Id": userId
                        }
                    });
                    if (!res.ok) throw new Error();
                    alert("Pedido excluído!");
                    carregarPedidos();
                } catch {
                    alert("Erro ao excluir pedido.");
                }
            });

            tbody.appendChild(row);
        });

    } catch (error) {
        console.error("Erro ao carregar pedidos:", error);
        mensagemErro.style.display = "block";
    }
}
