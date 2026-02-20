const state = {
  settings: null,
  toastTimer: null,
  loadedColumns: [],
  loadedRows: [],
  currentListType: 'acabamento',
  loadedKpis: null,
  dynamicEfficiencyEnabled: false,
  users: [],
  userPassTarget: null,
  adminSettingsPassword: null,
  isLoginsSectionUnlocked: false,
  adminPendingAction: null,
  authUser: null,
  needsFirstSetup: false,
  fiscalUser: null,
  fiscalPreviewItems: [],
  fiscalNfs: [],
  fiscalNfItems: [],
  fiscalSelectionNf: null,
  fiscalView: 'menu',
  fiscalHistoryRows: [],
  fiscalConfirmMode: null,
  fiscalDeleteReason: '',
  fiscalComprasTab: 'consulta',
  fiscalComprasRows: [],
  fiscalComprasFilters: {},
  fiscalComprasFornecedores: [],
  fiscalComprasCan: {},
  pcpSummary: null,
  pcpEfficiencySnapshot: null,
  pcpDashboardSnapshot: null,
  pcpMoldagemSnapshot: null,
  pcpEfficiencySelectedMachines: [],
  pcpEfficiencyHourMode: 'auto',
  pcpView: 'menu',
  pcpSelection: {
    freight: null,
    client: null,
    order: null
  },
  rhView: 'menu',
  rhHeSubView: 'dashboard',
  rhContext: null,
  rhSetores: [],
  rhColaboradores: [],
  rhHorasExtras: [],
  rhKpis: null,
  rhUsers: [],
  rhAudit: [],
  rhPendingConfirm: null,
  rhAtrasosKpis: null,
  rhAtrasosRows: [],
  rhSolicHeTab: 'form',
  rhSolicHeRows: [],
  rhSolicHeFilters: {},
  rhImportWorkbookPath: ''
};

const MOLDAGEM_HIDDEN_COLUMNS = new Set(['BUCHA PRENSADAS', 'PESO UTILIZADO']);
const MOLDAGEM_CENTER_COLUMNS = new Set(['COD', 'QTD', 'PC', 'REQ', 'CLIENTE']);
const MOLDAGEM_STATUS_PRENSADO = 'PRENSADO';
const MOLDAGEM_STATUS_ESTOQUE = 'BUCHA ESTOQUE';
const RNC_SECTOR_COLOR_MAP = {
  PCP: '#16a34a',
  PROJETO: '#facc15',
  USINAGEM: '#0ea5e9',
  ACABAMENTO: '#f97316',
  MOLDAGEM: '#1e40af',
  FORNECEDOR: '#86efac'
};
const STANDARD_CENTER_COLUMNS = new Set(['LINHA', 'QTD', 'N PC', 'N REQ', 'CLIENTE', 'PROGRAMA', 'BUCHA', 'ENTREGA', 'OBS']);
const MACHINE_LIST_TYPES = new Set(['fagor1', 'fagor2', 'mcs1', 'mcs2', 'mcs3']);
const READY_OBS_VALUES = new Set(['PRONTO', 'PALOMA', 'DARCI', 'REBECA', 'GUSTAVO', 'GUSTAVO S']);
const PCP_EFF_OP_BASED_TYPES = new Set(['fagor1', 'mcs1']);
const PCP_STATUS_IMAGE_PATHS = {
  rocket: '../../img/pcp-status-rocket.svg',
  okay: '../../img/pcp-status-okay.svg',
  alert: '../../img/pcp-status-alert.svg'
};
const MATERIAL_COLOR_MAP = {
  'Puro': '#2b6ef2',
  'Bronze Free': '#f7b267',
  'Grafite': '#c2c9d1',
  'T-46': '#78d18b',
  'Carbono': '#7a7f87',
  'Molibdenio': '#4f565f',
  'Fibra de vidro azul': '#1f4a8a',
  'Bronze Low': '#d97706',
  'Fibra de vidro branca': '#dff5ff',
  'Fibra de vidro amarela': '#f2d53c',
  'Fibra de vidro preta': '#111111'
};
const CNC_MACHINE_COLOR_MAP = {
  'FAGOR 1': '#2563eb',
  'FAGOR 2': '#1e3a8a',
  'MCS 1': '#86efac',
  'MCS 2': '#f97316',
  'MCS 3': '#8b5a2b'
};
const SETTINGS_UNLOCK_PASSWORD = '0604';
const FEATURE_LOG_GROUPS = [
  {
    version: '1.6.4',
    title: 'Backup - Restauração Segura e UX',
    date: '19/02/2026',
    items: [
      'Backup manual: botão "Executar backup agora" ajustado para forçar execução mesmo quando já houve backup automático no dia.',
      'Configuração: layout da área de backup reorganizado para visual mais limpo, moderno e corporativo.',
      'Configuração: botão de backup agora com feedback de execução (estado carregando, texto dinâmico e bloqueio de clique duplo).',
      'Novo botão "Usar backup" com fluxo de restauração no sistema.',
      'Restauração: modal com seleção de backup disponível e confirmação obrigatória da senha administrativa 3 vezes.',
      'Segurança: validação das 3 senhas feita também no backend antes de restaurar arquivos.',
      'Privacidade: referência visual da senha administrativa removida dos placeholders do modal.'
    ]
  },
  {
    version: '1.6.3',
    title: 'Configuração - Backup Automático Diário',
    date: '19/02/2026',
    items: [
      'Configuração: backup automático vinculado à Pasta raiz, sem necessidade de selecionar pasta manualmente.',
      'Backup salvo automaticamente na subpasta "Backup" dentro da Pasta raiz.',
      'Backup automático diário às 08:00 adicionado no app, com geração de pasta própria por execução.',
      'Regra diária global: apenas 1 backup por dia, independente de quem abrir o app.',
      'Retenção automática: backups com mais de 15 dias são removidos para manter organização e espaço.',
      'Novo botão "Executar backup agora" para disparo manual sob demanda.',
      'Backups incluem configurações e bases críticas (Fiscal Compras, RH e arquivos fiscais disponíveis).'
    ]
  },
  {
    version: '1.6.2',
    title: 'Menu - Atalhos de Teclado e Navegação Rápida',
    date: '19/02/2026',
    items: [
      'Menu inicial: atalhos de teclado adicionados para acesso rápido às áreas principais.',
      'Atalhos: tecla 1 (Listagem), 2 (PCP), 3 (Fiscal), 4 (RH), 5 (Configuração) e 6 (LOG).',
      'Atalhos respeitam permissões do usuário e bloqueiam acesso quando a área estiver desabilitada.',
      'Cards do Menu exibem indicador visual do atalho (badge numérico) para facilitar memorização e uso diário.',
      'Navegação rápida habilitada apenas quando o painel Menu está ativo, evitando conflito com campos de digitação.'
    ]
  },
  {
    version: '1.6.1',
    title: 'Menu Inicial Profissional (UX + Navegação)',
    date: '19/02/2026',
    items: [
      'Nova aba "Menu" adicionada como tela inicial do app (em vez de abrir direto na Listagem).',
      'Novo painel de navegação com cards para acessar: Listagem, PCP, Fiscal, RH, Configuração e LOG.',
      'Pós-login ajustado para abrir automaticamente o Menu inicial.',
      'Cards do Menu com validação de permissões por usuário (áreas sem acesso ficam desabilitadas).',
      'Visual do Menu redesenhado com estilo corporativo: hero inicial, grid responsivo e hierarquia tipográfica mais didática.',
      'Cards enriquecidos com ícones por área, subtítulo contextual e CTA "Abrir área", melhorando leitura e usabilidade.',
      'Aba de navegação superior ajustada para suportar quebra de linha quando necessário, sem perder organização.'
    ]
  },
  {
    version: '1.6',
    title: 'Fiscal Compras - UX, Status e PDF Corporativo',
    date: '19/02/2026',
    items: [
      'Solicitacao de Compras: campo Observacoes no formulario movido para baixo do Total Geral Estimado, em largura total e sem redimensionamento manual.',
      'Solicitacao de Compras: layout do formulario ajustado para remover componentes encavalados e melhorar leitura em desktop e mobile.',
      'Solicitacao de Compras: status padronizados para PENDENTE (padrao ao gerar), APROVADA e CANCELADA.',
      'Compatibilidade: status antigos (CRIADA, PEDIDO_GERADO, FINALIZADA) convertidos automaticamente para o novo padrao.',
      'Solicitacao de Compras: correcao de calculo de subtotal/total para aceitar valores com virgula e ponto (ex: 13,05 e 13.05).',
      'Cotacoes: destaque visual da melhor opcao reforcado na tela (mais visivel para comparacao rapida).',
      'Consulta de Compras: filtros reorganizados em duas linhas com grid mais limpo para evitar encavalamento.',
      'Consulta de Compras: coluna Acoes redesenhada com estrutura mais clara (status, observacao de aprovacao e botoes).',
      'Aprovacao: observacao de aprovacao salva no banco (APROVACAO_OBS), visivel para todos e editavel apenas por usuario com permissao de aprovar compra.',
      'Consulta de Compras: status com badge visual e observacao de aprovacao com truncamento + tooltip para leitura completa.',
      'Fluxo: ao salvar ou salvar+PDF, a tela retorna automaticamente para a aba Consultar para confirmacao do usuario.',
      'PDF da Solicitacao de Compras reformulado com visual corporativo: cabecalho vermelho escuro com texto branco, layout mais limpo, Observacoes abaixo do Total Geral Estimado e destaque em verde na celula do melhor fornecedor.'
    ]
  },
  {
    version: '1.5.15',
    title: 'Fiscal - Menu',
    date: '18/02/2026',
    items: [
      'Fiscal: adicionado menu inicial (estilo RH) com atalho para abrir a listagem de NFs.',
      'Fiscal: botao "Voltar ao menu" na tela de NFs.'
    ]
  },
  {
    version: '1.5.14',
    title: 'RH - Lancamento Manual (HH:MM)',
    date: '18/02/2026',
    items: [
      'RH: campo de duracao no lancamento manual agora aceita HH:MM (ex: 1:30) em H.E e Atrasos.',
      'RH: exibicao de horas no RH padronizada para HH:MM.'
    ]
  },
  {
    version: '1.5.13',
    title: 'Hotfix - RH Solicitacao H.E',
    date: '18/02/2026',
    items: [
      'Hotfix: preload/Window API da aba Solicitacao H.E (corrige erro de funcao inexistente em algumas instalacoes).'
    ]
  },
  {
    version: '1.5.12',
    title: 'RH - Solicitacao H.E (Compatibilidade)',
    date: '18/02/2026',
    items: [
      'RH: mensagem amigavel quando a maquina estiver com versao antiga (evita erro "rhSolicHeList is not a function").',
      'RH: orientacao para atualizar/reinstalar o Setup quando faltar API da aba Solicitação H.E.'
    ]
  },
  {
    version: '1.5.11',
    title: 'RH - Solicitacao H.E (Layout + Total HH:MM)',
    date: '18/02/2026',
    items: [
      'RH: Solicitacao de Horas Extras com layout mais limpo (campos alinhados e finalidade em largura total).',
      'RH: campo Finalidade com tamanho fixo (sem redimensionar).',
      'RH: Solicitante e Data da solicitacao automaticos (usuario logado + momento da geracao).',
      'RH: Total de horas agora em formato HH:MM (ex: 1:30) na tela e na consulta.',
      'PDF: Total de horas tambem em formato HH:MM.'
    ]
  },
  {
    version: '1.5.10',
    title: 'RH (Atrasos) + Importacao Melhorada + Atalho de Config no Login',
    date: '18/02/2026',
    items: [
      'Login: botao "Configuracao" e "Caminho Fiscal" direto na tela de acesso para facilitar o primeiro setup.',
      'Login: aviso mais claro quando o Banco Pedidos nao for encontrado (orienta configurar Caminho Fiscal R:).',
      'RH: nova sub-aba "Dashboard atrasos" com KPIs por setor/colaborador e evolucao mensal.',
      'RH: importacao agora traz H.E (S2) e Atrasos (R2) automaticamente por aba/colaborador.',
      'RH: modal "Lancar H.E" agora permite escolher entre lancar H.E ou Atraso no mesmo lugar.',
      'RH: validacao anti-duplicidade (atualiza registro existente por colaborador/data/tipo para evitar duplicados).'
    ]
  },
  {
    version: '1.5.9',
    title: 'RNC Completo + Rastreio SEDEX + Ajustes de Moldagem',
    date: '16/02/2026',
    items: [
      'Novo Relatorio RNC no Dashboard (pagina dedicada) com KPIs: Qntd RNC, Custo RNC, Custo medio, Tempo perdido, Setor com mais RNC, Setor com maior custo e Maior motivo.',
      'RNC: leitura direta do banco (Relatorio de Nao Conformidade BANCO DE DADOS.xlsm) usando Caminho Fiscal (R:) e aba RNC, tolerante a planilha aberta no Excel.',
      'RNC: calculo por data no formato pt-BR (ex: 12/02/26) e suporte a Segs Perdidos como HH:MM:SS ou estimativa 55s * Qntd quando vazio/zero.',
      'Graficos RNC: pizza % por setor com cores fixas, barras por setor, por motivo (quantidade e custo), tempo perdido por maquina (somente USINAGEM) e quantidade por material.',
      'Usinagem CNC: novos KPIs com Pecas Refugadas (RNC/USINAGEM) e Tempo Perdido (RNC/USINAGEM) e reorganizacao dos KPIs em 2 linhas.',
      'CNC: Percentual de contribuicao ajustado para ocupar metade da largura, liberando o lado para Custo apurado por maquina (RNC/USINAGEM).',
      'Moldagem: Buchas por Prensa e OP por Prensa corrigidos (apenas Prensa 1/2/3), lendo pela Tabela9 (cabecalho na linha 15) e consolidando resultados.',
      'Moldagem: KPI "OPs por Material" renomeado para "Ops Prensadas" na pagina de processamento.',
      'Dashboard PCP: caminho da planilha diaria atualizado para estrutura por Ano/Mes dentro de PROGRAMAÇÃO DE PRODUÇÃO-DIARIA.',
      'FISCAL: consulta de rastreio SEDEX com validacao do codigo e exibicao de status/destino via API (SeuRastreio).',
      'Configuracao: campo para salvar chave da API de rastreio e uso no botao Consultar dentro do FISCAL.',
      'RH: nova aba com ferramenta Gestao de Horas Extras (H.E) com filtros, dashboard, lancamentos e exportacao (Excel/PDF).',
      'RH: permissoes separadas em Configuracao (Logins) -> RH (visualizar) e RH Editar (lancar/editar/apagar + configuracoes).',
      'RH: KPIs por colaborador/setor e margem por faixas (25h=50%, 40h=60%, 60h=80%, acima=100%).'
    ]
  },
  {
    version: '1.5.8',
    title: 'Dashboard PCP 2.0 - PDF Profissional, Moldagem 2 e CNC Avancado',
    date: '15/02/2026',
    items: [
      'Dashboard renomeado e consolidado com identidade visual atual (Gestao Unificada / Parts Seals) e ajustes de abas.',
      'Visao Geral reorganizada com graficos verticais, melhor distribuicao de espaco e leitura mais limpa em tela.',
      'Acabamento por operador ajustado para ler nomes na coluna M (Acabamento) e somar pecas pela coluna QTD.',
      'Calculo de Pecas Planejadas corrigido pela regra da Tabela7 (detecao de fim por linhas vazias consecutivas).',
      'Exportacao PDF totalmente reformulada em tema claro e padrao A4 paisagem, independente da resolucao do monitor.',
      'PDF com cabecalhos profissionais por pagina (logo + secao + dia da semana + data) em Visao Geral, Moldagem, Moldagem 2 e CNC.',
      'Quebras de pagina inteligentes para evitar cortes de cards/graficos e preservar a leitura por secao.',
      'Moldagem criada como segunda secao do Dashboard (na mesma ferramenta), com KPIs e graficos dedicados.',
      'Leitura da aba Moldagem padronizada por colunas fixas: A material, C estoque, D buchas, E utilizado, F refugo, G OP.',
      'KPIs da Moldagem corrigidos: Material de Maior Saida por material utilizado e Material de Mais Buchas por quantidade de buchas.',
      'Moldagem 2 adicionada com nova pagina no PDF: Quantidade de Material Processado (Kg), Quantidade de Buchas por Material e 3 graficos de eficiencia por operador.',
      'Titulos da secao Moldagem padronizados: primeira pagina como Moldagem - Visao Geral e segunda pagina como Moldagem - Processamento.',
      'Novos KPIs executivos na Moldagem 2: OPs por Material, Media KG por OP, Media Buchas por OP, Indice de Refugo, Material com Maior Refugo, Participacao do Material Lider, Buchas por Kg e Top Operador por Peso.',
      'Usinagem Maquinas CNC adicionada como terceira secao do Dashboard com KPIs por maquina (FAGOR 1/2, MCS 1/2/3).',
      'CNC com graficos comparativos verticais (Produzidas x Planejadas) para pecas e OPs na mesma linha, com cores fixas por maquina.',
      'CNC com grafico de Percentual de Contribuicao (pizza), legenda ajustada no PDF e exibicao de percentuais dentro das fatias e em callouts para fatias pequenas.',
      'KPIs CNC atualizados com foco operacional: Gap de Pecas (Planejadas - Produzidas) e Gap de OPs (Planejadas - Produzidas).',
      'Scrollbars estilizadas em tema premium (trilho escuro + thumb elegante), removendo o aspecto branco padrao do sistema.',
      'Aplicativo configurado para abrir maximizado por padrao.'
    ]
  },
  {
    version: '1.5.7',
    title: 'Dashboard PCP - Correcao de Leitura e KPIs',
    date: '14/02/2026',
    items: [
      'Dashboard PCP: leitura da planilha reforcada para funcionar mesmo com o arquivo aberto no Excel (fallback por buffer/copia temporaria).',
      'Dashboard PCP: selecao de aba ajustada para reconhecer Planilha1 e Planilha 1 (alem de Tabela7).',
      'Dashboard PCP: deteccao do cabecalho da Tabela7 ampliada para varrer todas as abas e escolher automaticamente a correta.',
      'Dashboard PCP: cabecalho dinamico mais tolerante a deslocamento de colunas e variacoes de estrutura.',
      'KPIs corrigidos: OPs Planejadas contam linhas validas de QTD; OPs Feitas contam somente linhas com OBS = PRONTO.',
      'KPIs corrigidos: Pecas Feitas continuam sendo soma do QTD apenas quando OBS = PRONTO.'
    ]
  },
  {
    version: '1.5.6',
    title: 'Correcao de Baixas Duplicadas no PCP Acompanhamento',
    date: '14/02/2026',
    items: [
      'PCP Acompanhamento: ajuste do cruzamento para evitar que baixa de uma linha contamine outra linha igual com mesma OP (N REQ).',
      'PCP Acompanhamento: priorizacao de validacao por LINHA em cenarios de duplicidade de itens.',
      'PCP Acompanhamento: fallback sem LINHA somente quando o item for realmente unico.',
      'PCP Acompanhamento: status agora mostra detalhe do valor da celula para maquina/projeto (ex: BAIXA MAQUINA - NBR).',
      'Ajustes de consistencia no matching de Acabamento, Maquina e Projeto no resumo de fretes.'
    ]
  },
  {
    version: '1.5.5',
    title: 'Login Global, Permissoes Dinamicas e AutoComplete Fiscal',
    date: '14/02/2026',
    items: [
      'Novo login inicial do sistema com usuario/senha ao abrir o app.',
      'Sessao do usuario logado reaproveitada no Fiscal (sem reinserir senha para baixar/editar).',
      'Permissoes por area agora dinamicas (nao apenas PCP/FISCAL), preparadas para novas abas futuras.',
      'Cadastro de usuarios atualizado com permissoes por area geradas automaticamente.',
      'Tabela de logins em Configuracao atualizada com controles visuais mais modernos (switch estilo iOS).',
      'Secao de logins em Configuracao mantida oculta e liberada somente com senha administrativa 0604.',
      'Campo Cliente no cadastro de NF com sugestoes inteligentes de clientes ja cadastrados.',
      'Busca de cliente com tolerancia a acentos e variacao de maiusculas/minusculas.'
    ]
  },
  {
    version: '1.5.4',
    title: 'Nova Aba Fiscal e Cadastro de NF',
    date: '13/02/2026',
    items: [
      'Nova aba FISCAL com acesso por usuario/senha (ADM/1234) solicitado em toda abertura.',
      'Configuracao: campo Caminho Fiscal (R:) para localizar banco Pedidos.xlsx e BancoDeOrdens.xlsm.',
      'Fiscal: botao Cadastrar NF com busca de itens por pedidos/requisicoes e gravacao no banco de pedidos.',
      'Fiscal: lista de NFs emitidas com drilldown para visualizar itens da nota.',
      'Fiscal: edicao de NF (status/rastreio/data despache) com confirmacao de login e registro de quem alterou (BAIXADO POR + HORARIO).',
      'Fiscal: listagem de NFs exibe pedidos contemplados, rastreio e data de despache.',
      'Fiscal: botao Pesquisar NF para abrir/editar rapidamente pelo numero da nota.',
      'Fiscal: botao Apagar NF (com motivo + confirmacao de login) e historico salvo no banco (FISCAL_HISTORICO / FISCAL_NF_APAGADAS).',
      'Fiscal: botao Historico para visualizar NFs apagadas (NF, usuario, data/hora e motivo).',
      'Fiscal: listagem de NFs agora exibe requisicoes por pedido; pesquisa de NF aceita NF, pedido ou requisicao.',
      'Configuracao: cadastro de logins (planilha compartilhada) com permissao de acesso ao FISCAL.'
    ]
  },
  {
    version: '1.5.3',
    title: 'Exportacao Inteligente do Painel de Eficiencia',
    date: '13/02/2026',
    items: [
      'Eficiencia: botao para exportar o painel em imagem PNG.',
      'Eficiencia: filtro por maquinas no painel com selecao por checkbox.',
      'Exportacao respeita as maquinas selecionadas no momento da captura.',
      'Captura automatica da area completa (visao geral + cards de maquinas) sem depender do viewport.',
      'Ajustes de enquadramento para evitar corte lateral e inferior na imagem final.',
      'Calibracao fina de margens de exportacao para reduzir excesso de area em branco.',
      'PCP Acompanhamento: visao inicial de fretes agora lista clientes e quantidade de pedidos por tipo de frete.',
      'Rodape: creditos atualizados para Desenvolvido por KuruJoss com exibicao da logo no final da pagina.',
      'Aplicativo: icone configurado para janela/build (suporte a KuruJossIcon.png e KuruJossIcon.ico).'
    ]
  },
  {
    version: '1.5.2',
    title: 'PCP Operacional e Painel de Eficiencia',
    date: '13/02/2026',
    items: [
      'Aba PCP protegida por modal de senha em toda abertura.',
      'PCP reorganizado com menu de ferramentas.',
      'Nova ferramenta Acompanhamento com resumo por frete e drilldown: cliente > pedido > itens.',
      'Acompanhamento: status por item em cores (acabamento, maquina, projeto, sem baixa).',
      'Acompanhamento: pedidos consolidados sem repeticao e exibicao da OP (N REQ) no detalhe.',
      'Nova ferramenta Eficiencia com leitura das maquinas do dia e visao geral.',
      'Eficiencia: emissao por horario dinamico ou horarios fixos (10:00, 12:00, 15:00).',
      'Eficiencia: grafico bateria para ordens e pecas (atual x planejado no horario).',
      'Eficiencia: graficos de coluna por maquina e na visao geral (feitas x planejadas no horario).',
      'Eficiencia: status visual com foguete, joinha e bomba conforme desempenho.',
      'Eficiencia: seletor de data proprio para emitir em dias diferentes.',
      'Eficiencia: filtro por maquina para escolher quais cards aparecem no painel.',
      'Eficiencia: exportacao do painel em imagem PNG respeitando as maquinas selecionadas.',
      'Quando nenhuma data e selecionada, o sistema usa a data atual automaticamente.'
    ]
  },
  {
    version: '1.5.1',
    title: 'KPIs Avancados e Regras de Criticidade',
    date: '12/02/2026',
    items: [
      'KPIs adicionados na listagem Projeto: Total de Programas, Programas Novos e Programas Otimizados.',
      'KPI de itens com PROJETO LIBERADO baseado na coluna OBS.',
      'Grafico em bateria da eficiencia de liberacao (liberados x faltantes).',
      'Novo switch Ordens faltantes na tela de filtros (visivel somente na listagem Projeto).',
      'Filtro Ordens faltantes mostra apenas ordens sem baixa de projeto.',
      'Ordens com OBS = PRONTO passam a contar como PROJETO LIBERADO.',
      'Quando OBS = PRONTO, a tabela exibe PROJETO LIBERADO e destaca a linha em verde.',
      'Acabamento: KPIs de ordens/pecas criticas e filtro dedicado de Ordens criticas.',
      'Acabamento: segregacao por operador com ordens, pecas, percentuais e quantidades criticas.',
      'Acabamento: N REQ duplicado passa a classificar a ordem como critica.',
      'Ordens faltantes e Ordens criticas habilitados para todas as listagens.',
      'Mestra: KPIs de ordens/pecas criticas e percentuais adicionados.',
      'Moldagem: filtro Ordens criticas oculto (permanece apenas Ordens faltantes).',
      'Moldagem: novos KPIs de peso processado total, ordens processadas e buchas processadas com STATUS = PRENSADO.',
      'Moldagem: o filtro de Ordens faltantes ignora status BUCHA ESTOQUE (nao conta como pendente).'
    ]
  },
  {
    version: '1.0.0',
    title: 'Inicio de Tudo',
    date: '10/02/2026',
    items: [
      'Selecao de listagem por dia (campo de data).',
      'Selecao de tipo de listagem: Acabamento, Projeto, FAGOR 1/2, MCS 1/2/3, Moldagem e Mestra.',
      'Carregamento automatico da planilha com base em data + tipo de lista.',
      'Filtro por coluna e valor digitado.',
      'Funcao para filtrar ordens de moldagem (toggle).',
      'Botao para limpar filtros e restaurar a tabela completa.',
      'Tabela com cabecalho fixo e rolagem vertical.',
      'KPIs de producao (ordens, pecas e eficiencias).',
      'Indicadores visuais em bateria para desempenho.',
      'Eficiencia dinamica para listas de maquinas (com hora opcional).',
      'Destaques visuais por status e por regras de negocio.',
      'Selecao e salvamento da pasta raiz de trabalho.',
      'Salvamento de configuracoes do sistema.',
      'Alternancia de tema claro/escuro.',
      'Troca de logo conforme tema.',
      'Notificacoes na tela (toast) para sucesso/erro.',
      'Abas de navegacao: Listagem, Configuracao e LOG Atualizacoes.'
    ]
  },
  {
    version: '1.5.0',
    title: 'Implementacao da Lista Mestra',
    date: '11/02/2026',
    items: [
      'Listagem Mestra consolidada (Projeto + Maquinas + Acabamento).'
    ]
  }
];

function ensureRhSolicHeApiAvailable() {
  const api = window.api || {};
  const required = [
    'rhSolicHeUpsert',
    'rhSolicHeList',
    'rhSolicHeDelete',
    'rhSolicHePdfRegen',
    'rhSolicHePdfOpen'
  ];
  const missing = required.filter((key) => typeof api[key] !== 'function');
  if (!missing.length) {
    return;
  }
  throw new Error('A aba "Solicitação de H.E" precisa da versão mais nova do app. Atualize/reinstale o Setup (>= 1.5.13).');
}

const elements = {
  body: document.body,
  tabMenu: document.getElementById('tab-menu'),
  tabMain: document.getElementById('tab-main'),
  tabPcp: document.getElementById('tab-pcp'),
  tabFiscal: document.getElementById('tab-fiscal'),
  tabRh: document.getElementById('tab-rh'),
  tabSettings: document.getElementById('tab-settings'),
  tabLog: document.getElementById('tab-log'),
  panelMenu: document.getElementById('panel-menu'),
  panelMain: document.getElementById('panel-main'),
  panelPcp: document.getElementById('panel-pcp'),
  panelFiscal: document.getElementById('panel-fiscal'),
  panelRh: document.getElementById('panel-rh'),
  panelSettings: document.getElementById('panel-settings'),
  panelLog: document.getElementById('panel-log'),
  menuOpenMain: document.getElementById('menu-open-main'),
  menuOpenPcp: document.getElementById('menu-open-pcp'),
  menuOpenFiscal: document.getElementById('menu-open-fiscal'),
  menuOpenRh: document.getElementById('menu-open-rh'),
  menuOpenSettings: document.getElementById('menu-open-settings'),
  menuOpenLog: document.getElementById('menu-open-log'),
  brandLogo: document.getElementById('brand-logo'),
  themeSwitch: document.getElementById('theme-switch'),
  rootFolderInput: document.getElementById('root-folder-input'),
  btnSelectFolder: document.getElementById('btn-select-folder'),
  fiscalRootInput: document.getElementById('fiscal-root-input'),
  btnSelectFiscalRoot: document.getElementById('btn-select-fiscal-root'),
  trackingApiKeyInput: document.getElementById('tracking-api-key-input'),
  btnRunBackupNow: document.getElementById('btn-run-backup-now'),
  btnRestoreBackup: document.getElementById('btn-restore-backup'),
  btnSaveSettings: document.getElementById('btn-save-settings'),
  btnUnlockLogins: document.getElementById('btn-unlock-logins'),
  loginsLockedNote: document.getElementById('logins-locked-note'),
  loginsAdminSection: document.getElementById('logins-admin-section'),
  btnUserAdd: document.getElementById('btn-user-add'),
  usersEmpty: document.getElementById('users-empty'),
  usersTableHead: document.getElementById('users-table-head'),
  usersTableBody: document.getElementById('users-table-body'),
  dateInput: document.getElementById('date-input'),
  listTypeSelect: document.getElementById('list-type-select'),
  btnLoad: document.getElementById('btn-load'),
  filterColumnSelect: document.getElementById('filter-column-select'),
  filterValueInput: document.getElementById('filter-value-input'),
  btnApplyFilter: document.getElementById('btn-apply-filter'),
  btnClearFilter: document.getElementById('btn-clear-filter'),
  filterBuchaNotEmpty: document.getElementById('filter-bucha-not-empty'),
  filterProjectMissing: document.getElementById('filter-project-missing'),
  filterAcabamentoCritical: document.getElementById('filter-acabamento-critical'),
  dynamicEffSwitch: document.getElementById('dynamic-eff-switch'),
  dynamicHourInput: document.getElementById('dynamic-hour-input'),
  kpiSection: document.getElementById('kpi-section'),
  resultPath: document.getElementById('result-path'),
  tableHead: document.querySelector('#result-table thead'),
  tableBody: document.querySelector('#result-table tbody'),
  logUpdatesContent: document.getElementById('log-updates-content'),
  footerVersion: document.getElementById('footer-version'),
  toast: document.getElementById('toast'),
  pcpMenuView: document.getElementById('pcp-menu-view'),
  pcpAcompView: document.getElementById('pcp-acomp-view'),
  pcpEffView: document.getElementById('pcp-eff-view'),
  pcpOpenAcomp: document.getElementById('pcp-open-acomp'),
  pcpOpenEff: document.getElementById('pcp-open-eff'),
  pcpOpenDashboard: document.getElementById('pcp-open-dashboard'),
  pcpBackMenu: document.getElementById('pcp-back-menu'),
  pcpEffBackMenu: document.getElementById('pcp-eff-back-menu'),
  pcpDashboardView: document.getElementById('pcp-dashboard-view'),
  pcpDashboardBackMenu: document.getElementById('pcp-dashboard-back-menu'),
  pcpDashboardRefresh: document.getElementById('pcp-dashboard-refresh'),
  pcpDashboardExportPdf: document.getElementById('pcp-dashboard-export-pdf'),
  pcpDashboardDate: document.getElementById('pcp-dashboard-date'),
  pcpDashboardTitle: document.getElementById('pcp-dashboard-title'),
  pcpDashboardPdfHeader: document.getElementById('pcp-dashboard-pdf-header'),
  pcpDashboardPdfTitle: document.getElementById('pcp-dashboard-pdf-title'),
  pcpDashboardEmpty: document.getElementById('pcp-dashboard-empty'),
  pcpDashboardOverview: document.getElementById('pcp-dashboard-overview'),
  pcpDashboardCharts: document.getElementById('pcp-dashboard-charts'),
  pcpDashboardExportArea: document.getElementById('pcp-dashboard-export-area'),
  pcpMoldagemSection: document.getElementById('pcp-moldagem-section'),
  pcpMoldagemTitle: document.getElementById('pcp-moldagem-title'),
  pcpMoldagemEmpty: document.getElementById('pcp-moldagem-empty'),
  pcpMoldagemOverview: document.getElementById('pcp-moldagem-overview'),
  pcpMoldagemCharts: document.getElementById('pcp-moldagem-charts'),
  pcpCncSection: document.getElementById('pcp-cnc-section'),
  pcpCncTitle: document.getElementById('pcp-cnc-title'),
  pcpCncEmpty: document.getElementById('pcp-cnc-empty'),
  pcpCncOverview: document.getElementById('pcp-cnc-overview'),
  pcpCncCharts: document.getElementById('pcp-cnc-charts'),
  pcpRncSection: document.getElementById('pcp-rnc-section'),
  pcpRncTitle: document.getElementById('pcp-rnc-title'),
  pcpRncEmpty: document.getElementById('pcp-rnc-empty'),
  pcpRncOverview: document.getElementById('pcp-rnc-overview'),
  pcpRncCharts: document.getElementById('pcp-rnc-charts'),
  pcpEffRefresh: document.getElementById('pcp-eff-refresh'),
  pcpEffExport: document.getElementById('pcp-eff-export'),
  pcpEffDate: document.getElementById('pcp-eff-date'),
  pcpEffTitle: document.getElementById('pcp-eff-title'),
  pcpEffEmpty: document.getElementById('pcp-eff-empty'),
  pcpEffMachineFilter: document.getElementById('pcp-eff-machine-filter'),
  pcpEffExportArea: document.getElementById('pcp-eff-export-area'),
  pcpEffOverall: document.getElementById('pcp-eff-overall'),
  pcpEffGrid: document.getElementById('pcp-eff-grid'),
  pcpEffTimeAuto: document.getElementById('pcp-eff-time-auto'),
  pcpEffTime10: document.getElementById('pcp-eff-time-10'),
  pcpEffTime12: document.getElementById('pcp-eff-time-12'),
  pcpEffTime15: document.getElementById('pcp-eff-time-15'),
  pcpSubtitle: document.getElementById('pcp-subtitle'),
  pcpBreadcrumb: document.getElementById('pcp-breadcrumb'),
  pcpTableHead: document.getElementById('pcp-table-head'),
  pcpTableBody: document.getElementById('pcp-table-body'),
  pcpEmptyState: document.getElementById('pcp-empty-state'),
  appLoginModal: document.getElementById('app-login-modal'),
  appLoginUsername: document.getElementById('app-login-username'),
  appLoginPassword: document.getElementById('app-login-password'),
  appLoginError: document.getElementById('app-login-error'),
  appLoginConfig: document.getElementById('app-login-config'),
  appLoginFiscalRoot: document.getElementById('app-login-fiscal-root'),
  appLoginConfirm: document.getElementById('app-login-confirm'),
  fiscalMenuView: document.getElementById('fiscal-menu-view'),
  fiscalNfsView: document.getElementById('fiscal-nfs-view'),
  fiscalComprasView: document.getElementById('fiscal-compras-view'),
  fiscalOpenNfs: document.getElementById('fiscal-open-nfs'),
  fiscalOpenCompras: document.getElementById('fiscal-open-compras'),
  fiscalBackMenu: document.getElementById('fiscal-back-menu'),
  fiscalSubtitle: document.getElementById('fiscal-subtitle'),
  fiscalBtnCadastrarNf: document.getElementById('fiscal-btn-cadastrar-nf'),
  fiscalBtnPesquisarNf: document.getElementById('fiscal-btn-pesquisar-nf'),
  fiscalBtnPendencias: document.getElementById('fiscal-btn-pendencias'),
  fiscalBtnHistorico: document.getElementById('fiscal-btn-historico'),
  fiscalEmptyState: document.getElementById('fiscal-empty-state'),
  fiscalTableHead: document.getElementById('fiscal-table-head'),
  fiscalTableBody: document.getElementById('fiscal-table-body'),
  fiscalBtnVoltar: document.getElementById('fiscal-btn-voltar'),
  fiscalBtnEditarNf: document.getElementById('fiscal-btn-editar-nf'),
  fiscalBtnApagarNf: document.getElementById('fiscal-btn-apagar-nf'),
  fiscalTrackingCode: document.getElementById('fiscal-tracking-code'),
  fiscalTrackingCheck: document.getElementById('fiscal-tracking-check'),
  fiscalTrackingResult: document.getElementById('fiscal-tracking-result'),
  fiscalBreadcrumb: document.getElementById('fiscal-breadcrumb'),
  fiscalNfModal: document.getElementById('fiscal-nf-modal'),
  fiscalNfBackdrop: document.getElementById('fiscal-nf-backdrop'),
  fiscalNfNumber: document.getElementById('fiscal-nf-number'),
  fiscalNfClient: document.getElementById('fiscal-nf-client'),
  fiscalClientSuggestions: document.getElementById('fiscal-client-suggestions'),
  fiscalNfIdentifiers: document.getElementById('fiscal-nf-identifiers'),
  fiscalNfDatetime: document.getElementById('fiscal-nf-datetime'),
  fiscalNfError: document.getElementById('fiscal-nf-error'),
  fiscalNfCancel: document.getElementById('fiscal-nf-cancel'),
  fiscalNfSearch: document.getElementById('fiscal-nf-search'),
  fiscalNfConfirm: document.getElementById('fiscal-nf-confirm'),
  fiscalNfPreview: document.getElementById('fiscal-nf-preview'),
  fiscalNfPreviewTitle: document.getElementById('fiscal-nf-preview-title'),
  fiscalNfPreviewHead: document.getElementById('fiscal-nf-preview-head'),
  fiscalNfPreviewBody: document.getElementById('fiscal-nf-preview-body'),
  fiscalEditModal: document.getElementById('fiscal-edit-modal'),
  fiscalEditBackdrop: document.getElementById('fiscal-edit-backdrop'),
  fiscalEditStatus: document.getElementById('fiscal-edit-status'),
  fiscalEditRastreio: document.getElementById('fiscal-edit-rastreio'),
  fiscalEditDespache: document.getElementById('fiscal-edit-despache'),
  fiscalEditError: document.getElementById('fiscal-edit-error'),
  fiscalEditCancel: document.getElementById('fiscal-edit-cancel'),
  fiscalEditConfirm: document.getElementById('fiscal-edit-confirm'),
  fiscalConfirmModal: document.getElementById('fiscal-confirm-modal'),
  fiscalConfirmBackdrop: document.getElementById('fiscal-confirm-backdrop'),
  fiscalConfirmUserLabel: document.getElementById('fiscal-confirm-user-label'),
  fiscalConfirmError: document.getElementById('fiscal-confirm-error'),
  fiscalConfirmCancel: document.getElementById('fiscal-confirm-cancel'),
  fiscalConfirmOk: document.getElementById('fiscal-confirm-ok'),
  fiscalSearchModal: document.getElementById('fiscal-search-modal'),
  fiscalSearchBackdrop: document.getElementById('fiscal-search-backdrop'),
  fiscalSearchNf: document.getElementById('fiscal-search-nf'),
  fiscalSearchError: document.getElementById('fiscal-search-error'),
  fiscalSearchCancel: document.getElementById('fiscal-search-cancel'),
  fiscalSearchConfirm: document.getElementById('fiscal-search-confirm'),
  fiscalDeleteModal: document.getElementById('fiscal-delete-modal'),
  fiscalDeleteBackdrop: document.getElementById('fiscal-delete-backdrop'),
  fiscalDeleteReason: document.getElementById('fiscal-delete-reason'),
  fiscalDeleteError: document.getElementById('fiscal-delete-error'),
  fiscalDeleteCancel: document.getElementById('fiscal-delete-cancel'),
  fiscalDeleteConfirm: document.getElementById('fiscal-delete-confirm'),
  fiscalComprasBackMenu: document.getElementById('fiscal-compras-back-menu'),
  fiscalComprasRefresh: document.getElementById('fiscal-compras-refresh'),
  fiscalComprasNew: document.getElementById('fiscal-compras-new'),
  fiscalComprasSubtitle: document.getElementById('fiscal-compras-subtitle'),
  fiscalComprasTabForm: document.getElementById('fiscal-compras-tab-form'),
  fiscalComprasTabConsulta: document.getElementById('fiscal-compras-tab-consulta'),
  fiscalComprasViewForm: document.getElementById('fiscal-compras-view-form'),
  fiscalComprasViewConsulta: document.getElementById('fiscal-compras-view-consulta'),
  fiscalComprasId: document.getElementById('fiscal-compras-id'),
  fiscalComprasNumero: document.getElementById('fiscal-compras-numero'),
  fiscalComprasDataSolic: document.getElementById('fiscal-compras-data-solic'),
  fiscalComprasSolicitadoPor: document.getElementById('fiscal-compras-solicitado-por'),
  fiscalComprasSetor: document.getElementById('fiscal-compras-setor'),
  fiscalComprasDestinado: document.getElementById('fiscal-compras-destinado'),
  fiscalComprasStatus: document.getElementById('fiscal-compras-status'),
  fiscalComprasObservacoes: document.getElementById('fiscal-compras-observacoes'),
  fiscalComprasItemsWrap: document.getElementById('fiscal-compras-items-wrap'),
  fiscalComprasFornecedorList: document.getElementById('fiscal-compras-fornecedor-list'),
  fiscalComprasPagamentoList: document.getElementById('fiscal-compras-pagamento-list'),
  fiscalComprasTotalEstimado: document.getElementById('fiscal-compras-total-estimado'),
  fiscalComprasError: document.getElementById('fiscal-compras-error'),
  fiscalComprasAddItem: document.getElementById('fiscal-compras-add-item'),
  fiscalComprasCancel: document.getElementById('fiscal-compras-cancel'),
  fiscalComprasSave: document.getElementById('fiscal-compras-save'),
  fiscalComprasSavePdf: document.getElementById('fiscal-compras-save-pdf'),
  fiscalComprasFilterIni: document.getElementById('fiscal-compras-filter-ini'),
  fiscalComprasFilterFim: document.getElementById('fiscal-compras-filter-fim'),
  fiscalComprasFilterSolicitante: document.getElementById('fiscal-compras-filter-solicitante'),
  fiscalComprasFilterSetor: document.getElementById('fiscal-compras-filter-setor'),
  fiscalComprasFilterNumero: document.getElementById('fiscal-compras-filter-numero'),
  fiscalComprasFilterFornecedor: document.getElementById('fiscal-compras-filter-fornecedor'),
  fiscalComprasFilterItem: document.getElementById('fiscal-compras-filter-item'),
  fiscalComprasFilterStatus: document.getElementById('fiscal-compras-filter-status'),
  fiscalComprasFilterClear: document.getElementById('fiscal-compras-filter-clear'),
  fiscalComprasFilterApply: document.getElementById('fiscal-compras-filter-apply'),
  fiscalComprasConsultaError: document.getElementById('fiscal-compras-consulta-error'),
  fiscalComprasConsultaEmpty: document.getElementById('fiscal-compras-consulta-empty'),
  fiscalComprasTableHead: document.getElementById('fiscal-compras-table-head'),
  fiscalComprasTableBody: document.getElementById('fiscal-compras-table-body'),
  rhMenuView: document.getElementById('rh-menu-view'),
  rhOpenHe: document.getElementById('rh-open-he'),
  rhOpenSolicHe: document.getElementById('rh-open-solic-he'),
  rhHeView: document.getElementById('rh-he-view'),
  rhSolicHeView: document.getElementById('rh-solic-he-view'),
  rhHeBackMenu: document.getElementById('rh-he-back-menu'),
  rhHeRefresh: document.getElementById('rh-he-refresh'),
  rhHeAdd: document.getElementById('rh-he-add'),
  rhHeImport: document.getElementById('rh-he-import'),
  rhHeExportExcel: document.getElementById('rh-he-export-excel'),
  rhHeExportPdf: document.getElementById('rh-he-export-pdf'),
  rhSolicHeBackMenu: document.getElementById('rh-solic-he-back-menu'),
  rhSolicHeRefresh: document.getElementById('rh-solic-he-refresh'),
  rhSolicHeNew: document.getElementById('rh-solic-he-new'),
  rhSolicHeTabForm: document.getElementById('rh-solic-he-tab-form'),
  rhSolicHeTabConsulta: document.getElementById('rh-solic-he-tab-consulta'),
  rhSolicHeViewForm: document.getElementById('rh-solic-he-view-form'),
  rhSolicHeViewConsulta: document.getElementById('rh-solic-he-view-consulta'),
  rhSolicHeId: document.getElementById('rh-solic-he-id'),
  rhSolicHeNumero: document.getElementById('rh-solic-he-numero'),
  rhSolicHeDataHe: document.getElementById('rh-solic-he-datahe'),
  rhSolicHeColab: document.getElementById('rh-solic-he-colab'),
  rhSolicHeSetor: document.getElementById('rh-solic-he-setor'),
  rhSolicHeFinalidade: document.getElementById('rh-solic-he-finalidade'),
  rhSolicHeInicio: document.getElementById('rh-solic-he-inicio'),
  rhSolicHeFim: document.getElementById('rh-solic-he-fim'),
  rhSolicHeTotal: document.getElementById('rh-solic-he-total'),
  rhSolicHeSolicitante: document.getElementById('rh-solic-he-solicitante'),
  rhSolicHeDataSolic: document.getElementById('rh-solic-he-datasolic'),
  rhSolicHeError: document.getElementById('rh-solic-he-error'),
  rhSolicHeCancel: document.getElementById('rh-solic-he-cancel'),
  rhSolicHeSave: document.getElementById('rh-solic-he-save'),
  rhSolicHeFilterColab: document.getElementById('rh-solic-he-filter-colab'),
  rhSolicHeFilterSetor: document.getElementById('rh-solic-he-filter-setor'),
  rhSolicHeFilterIni: document.getElementById('rh-solic-he-filter-ini'),
  rhSolicHeFilterFim: document.getElementById('rh-solic-he-filter-fim'),
  rhSolicHeFilterSolicitante: document.getElementById('rh-solic-he-filter-solicitante'),
  rhSolicHeFilterDoc: document.getElementById('rh-solic-he-filter-doc'),
  rhSolicHeFilterClear: document.getElementById('rh-solic-he-filter-clear'),
  rhSolicHeFilterApply: document.getElementById('rh-solic-he-filter-apply'),
  rhSolicHeConsultaError: document.getElementById('rh-solic-he-consulta-error'),
  rhSolicHeConsultaEmpty: document.getElementById('rh-solic-he-consulta-empty'),
  rhSolicHeTable: document.getElementById('rh-solic-he-table'),
  rhSolicHeTableHead: document.getElementById('rh-solic-he-table-head'),
  rhSolicHeTableBody: document.getElementById('rh-solic-he-table-body'),
  rhFilterMonth: document.getElementById('rh-filter-month'),
  rhFilterSector: document.getElementById('rh-filter-sector'),
  rhFilterCollab: document.getElementById('rh-filter-collab'),
  rhFilterType: document.getElementById('rh-filter-type'),
  rhSubtabDashboard: document.getElementById('rh-subtab-dashboard'),
  rhSubtabAtrasos: document.getElementById('rh-subtab-atrasos'),
  rhSubtabRegistros: document.getElementById('rh-subtab-registros'),
  rhSubtabColaboradores: document.getElementById('rh-subtab-colaboradores'),
  rhSubtabSetores: document.getElementById('rh-subtab-setores'),
  rhSubtabConfig: document.getElementById('rh-subtab-config'),
  rhSubtabAudit: document.getElementById('rh-subtab-audit'),
  rhViewDashboard: document.getElementById('rh-view-dashboard'),
  rhViewAtrasos: document.getElementById('rh-view-atrasos'),
  rhViewRegistros: document.getElementById('rh-view-registros'),
  rhViewColaboradores: document.getElementById('rh-view-colaboradores'),
  rhViewSetores: document.getElementById('rh-view-setores'),
  rhViewConfig: document.getElementById('rh-view-config'),
  rhViewAudit: document.getElementById('rh-view-audit'),
  rhSectorModal: document.getElementById('rh-sector-modal'),
  rhSectorBackdrop: document.getElementById('rh-sector-backdrop'),
  rhSectorId: document.getElementById('rh-sector-id'),
  rhSectorNome: document.getElementById('rh-sector-nome'),
  rhSectorGestor: document.getElementById('rh-sector-gestor'),
  rhSectorCc: document.getElementById('rh-sector-cc'),
  rhSectorError: document.getElementById('rh-sector-error'),
  rhSectorCancel: document.getElementById('rh-sector-cancel'),
  rhSectorConfirm: document.getElementById('rh-sector-confirm'),
  rhCollabModal: document.getElementById('rh-collab-modal'),
  rhCollabBackdrop: document.getElementById('rh-collab-backdrop'),
  rhCollabId: document.getElementById('rh-collab-id'),
  rhCollabNome: document.getElementById('rh-collab-nome'),
  rhCollabMatricula: document.getElementById('rh-collab-matricula'),
  rhCollabSetor: document.getElementById('rh-collab-setor'),
  rhCollabCargo: document.getElementById('rh-collab-cargo'),
  rhCollabAdmissao: document.getElementById('rh-collab-admissao'),
  rhCollabStatus: document.getElementById('rh-collab-status'),
  rhCollabLimite: document.getElementById('rh-collab-limite'),
  rhCollabError: document.getElementById('rh-collab-error'),
  rhCollabCancel: document.getElementById('rh-collab-cancel'),
  rhCollabConfirm: document.getElementById('rh-collab-confirm'),
  rhHeModal: document.getElementById('rh-he-modal'),
  rhHeBackdrop: document.getElementById('rh-he-backdrop'),
  rhHeTitle: document.getElementById('rh-he-title'),
  rhHeModalSubtitle: document.getElementById('rh-he-modal-subtitle'),
  rhHeId: document.getElementById('rh-he-id'),
  rhHeKind: document.getElementById('rh-he-kind'),
  rhHeColaborador: document.getElementById('rh-he-colaborador'),
  rhHeData: document.getElementById('rh-he-data'),
  rhHeQtdLabel: document.getElementById('rh-he-qtd-label'),
  rhHeQtd: document.getElementById('rh-he-qtd'),
  rhHeTipoField: document.getElementById('rh-he-tipo-field'),
  rhHeTipo: document.getElementById('rh-he-tipo'),
  rhHeObs: document.getElementById('rh-he-obs'),
  rhHeJustField: document.getElementById('rh-he-just-field'),
  rhHeJust: document.getElementById('rh-he-just'),
  rhHeError: document.getElementById('rh-he-error'),
  rhHeCancel: document.getElementById('rh-he-cancel'),
  rhHeConfirm: document.getElementById('rh-he-confirm'),
  rhConfirmModal: document.getElementById('rh-confirm-modal'),
  rhConfirmBackdrop: document.getElementById('rh-confirm-backdrop'),
  rhConfirmMessage: document.getElementById('rh-confirm-message'),
  rhConfirmError: document.getElementById('rh-confirm-error'),
  rhConfirmCancel: document.getElementById('rh-confirm-cancel'),
  rhConfirmOk: document.getElementById('rh-confirm-ok'),
  rhImportModal: document.getElementById('rh-import-modal'),
  rhImportBackdrop: document.getElementById('rh-import-backdrop'),
  rhImportSource: document.getElementById('rh-import-source'),
  rhImportDate: document.getElementById('rh-import-date'),
  rhImportFile: document.getElementById('rh-import-file'),
  rhImportFilePicker: document.getElementById('rh-import-file-picker'),
  rhImportFilePath: document.getElementById('rh-import-file-path'),
  rhImportFilePick: document.getElementById('rh-import-file-pick'),
  rhImportError: document.getElementById('rh-import-error'),
  rhImportCancel: document.getElementById('rh-import-cancel'),
  rhImportConfirm: document.getElementById('rh-import-confirm'),
  userModal: document.getElementById('user-modal'),
  userModalBackdrop: document.getElementById('user-modal-backdrop'),
  userUsername: document.getElementById('user-username'),
  userPassword: document.getElementById('user-password'),
  userPermissionsWrap: document.getElementById('user-permissions-wrap'),
  userModalError: document.getElementById('user-modal-error'),
  userModalCancel: document.getElementById('user-modal-cancel'),
  userModalConfirm: document.getElementById('user-modal-confirm'),
  userPassModal: document.getElementById('user-pass-modal'),
  userPassBackdrop: document.getElementById('user-pass-backdrop'),
  userPassSubtitle: document.getElementById('user-pass-subtitle'),
  userPassInput: document.getElementById('user-pass-input'),
  userPassError: document.getElementById('user-pass-error'),
  userPassCancel: document.getElementById('user-pass-cancel'),
  userPassConfirm: document.getElementById('user-pass-confirm'),
  adminSettingsModal: document.getElementById('admin-settings-modal'),
  adminSettingsBackdrop: document.getElementById('admin-settings-backdrop'),
  adminSettingsPassword: document.getElementById('admin-settings-password'),
  adminSettingsError: document.getElementById('admin-settings-error'),
  adminSettingsCancel: document.getElementById('admin-settings-cancel'),
  adminSettingsConfirm: document.getElementById('admin-settings-confirm'),
  backupRestoreModal: document.getElementById('backup-restore-modal'),
  backupRestoreBackdrop: document.getElementById('backup-restore-backdrop'),
  backupRestoreSelect: document.getElementById('backup-restore-select'),
  backupRestorePass1: document.getElementById('backup-restore-pass1'),
  backupRestorePass2: document.getElementById('backup-restore-pass2'),
  backupRestorePass3: document.getElementById('backup-restore-pass3'),
  backupRestoreError: document.getElementById('backup-restore-error'),
  backupRestoreCancel: document.getElementById('backup-restore-cancel'),
  backupRestoreConfirm: document.getElementById('backup-restore-confirm')
};

function getManagedAreas() {
  const tabs = Array.from(document.querySelectorAll('.tabs .tab[data-tab]'));
  return tabs
    .map((tab) => ({
      key: String(tab.dataset.tab || '').trim().toLowerCase(),
      label: String(tab.textContent || '').trim() || String(tab.dataset.tab || '').trim().toUpperCase()
    }))
    .filter((area) => area.key && area.key !== 'main');
}

function getAreaDefaultPermission(areaKey) {
  return areaKey === 'fiscal' || areaKey === 'rh' || areaKey === 'rh_edit' ? false : true;
}

function getPermissionValue(permissions, areaKey) {
  if (permissions && Object.prototype.hasOwnProperty.call(permissions, areaKey)) {
    return !!permissions[areaKey];
  }
  return getAreaDefaultPermission(areaKey);
}

const FISCAL_COMPRAS_PERMISSION_DEFS = [
  { key: 'compras_view', title: 'Pode Visualizar Solicitação de Compras', subtitle: 'Permite abrir, consultar e ver histórico de compras.' },
  { key: 'compras_create', title: 'Pode Criar Solicitação', subtitle: 'Permite criar novas solicitações de compra.' },
  { key: 'compras_edit', title: 'Pode Editar Solicitação', subtitle: 'Permite editar solicitações de compra existentes.' },
  { key: 'compras_delete', title: 'Pode Excluir Solicitação', subtitle: 'Permite excluir solicitações de compra (com restrições de status).' },
  { key: 'compras_pdf', title: 'Pode Gerar PDF', subtitle: 'Permite gerar e abrir PDF profissional da solicitação.' },
  { key: 'compras_approve', title: 'Pode Aprovar Compra', subtitle: 'Permite aprovar solicitação e alterar status de aprovação.' }
];

function getPermissionValueWithFallback(permissions, key, fallbackValue = false) {
  if (permissions && Object.prototype.hasOwnProperty.call(permissions, key)) {
    return !!permissions[key];
  }
  return !!fallbackValue;
}

function collectUserModalPermissions() {
  const result = {};
  const inputs = Array.from(document.querySelectorAll('[data-user-permission-key]'));
  inputs.forEach((input) => {
    const key = String(input.getAttribute('data-user-permission-key') || '').trim().toLowerCase();
    if (!key) {
      return;
    }
    result[key] = !!input.checked;
  });
  return result;
}

function renderUserPermissionsEditor(permissionsSeed = {}) {
  if (!elements.userPermissionsWrap) {
    return;
  }

  elements.userPermissionsWrap.innerHTML = '';
  const managedAreas = getManagedAreas();
  managedAreas.forEach((area) => {
    const item = document.createElement('div');
    item.className = 'user-permission-item';

    const textWrap = document.createElement('div');
    textWrap.className = 'user-permission-text';
    const strong = document.createElement('strong');
    strong.textContent = `Acesso ${area.label}`;
    const subtitle = document.createElement('span');
    subtitle.textContent = `Permite abrir a área ${area.label}.`;
    textWrap.appendChild(strong);
    textWrap.appendChild(subtitle);

    const inputId = `user-can-${area.key}`;
    const label = document.createElement('label');
    label.className = 'switch-control switch-sm';
    label.setAttribute('for', inputId);

    const input = document.createElement('input');
    input.type = 'checkbox';
    input.id = inputId;
    input.setAttribute('data-user-permission-key', area.key);
    input.checked = getPermissionValue(permissionsSeed, area.key);

    const slider = document.createElement('span');
    slider.className = 'switch-slider';

    label.appendChild(input);
    label.appendChild(slider);
    item.appendChild(textWrap);
    item.appendChild(label);
    elements.userPermissionsWrap.appendChild(item);
  });

  // Permissão extra (não é uma aba): permite editar/excluir no RH.
  if (managedAreas.some((area) => area.key === 'rh')) {
    const item = document.createElement('div');
    item.className = 'user-permission-item';

    const textWrap = document.createElement('div');
    textWrap.className = 'user-permission-text';
    const strong = document.createElement('strong');
    strong.textContent = 'RH - Permissão de edição';
    const subtitle = document.createElement('span');
    subtitle.textContent = 'Permite lançar, editar e excluir registros de Horas Extras.';
    textWrap.appendChild(strong);
    textWrap.appendChild(subtitle);

    const inputId = 'user-can-rh-edit';
    const label = document.createElement('label');
    label.className = 'switch-control switch-sm';
    label.setAttribute('for', inputId);

    const input = document.createElement('input');
    input.type = 'checkbox';
    input.id = inputId;
    input.setAttribute('data-user-permission-key', 'rh_edit');
    input.checked = !!(permissionsSeed && permissionsSeed.rh_edit);

    const slider = document.createElement('span');
    slider.className = 'switch-slider';

    label.appendChild(input);
    label.appendChild(slider);
    item.appendChild(textWrap);
    item.appendChild(label);
    elements.userPermissionsWrap.appendChild(item);
  }

  if (managedAreas.some((area) => area.key === 'fiscal')) {
    FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
      const item = document.createElement('div');
      item.className = 'user-permission-item';

      const textWrap = document.createElement('div');
      textWrap.className = 'user-permission-text';
      const strong = document.createElement('strong');
      strong.textContent = `Fiscal - ${perm.title}`;
      const subtitle = document.createElement('span');
      subtitle.textContent = perm.subtitle;
      textWrap.appendChild(strong);
      textWrap.appendChild(subtitle);

      const inputId = `user-can-${perm.key}`;
      const label = document.createElement('label');
      label.className = 'switch-control switch-sm';
      label.setAttribute('for', inputId);

      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = inputId;
      input.setAttribute('data-user-permission-key', perm.key);
      input.checked = getPermissionValueWithFallback(permissionsSeed, perm.key, !!permissionsSeed.fiscal);

      const slider = document.createElement('span');
      slider.className = 'switch-slider';

      label.appendChild(input);
      label.appendChild(slider);
      item.appendChild(textWrap);
      item.appendChild(label);
      elements.userPermissionsWrap.appendChild(item);
    });
  }
}

function normalizeColumnName(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function normalizeForFilter(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}

function isMoldagem() {
  return state.currentListType === 'moldagem';
}

function isMachineList() {
  return MACHINE_LIST_TYPES.has(state.currentListType);
}

function isProjeto() {
  return state.currentListType === 'projeto';
}

function isAcabamento() {
  return state.currentListType === 'acabamento';
}

function isProjectReleasedObs(obsValue) {
  const normalizedObs = normalizeForFilter(obsValue);
  return normalizedObs.includes('PROJETO LIBERADO') || normalizedObs === 'PRONTO';
}

function isAcabamentoReadyObs(obsValue) {
  const normalizedObs = normalizeForFilter(obsValue);
  return READY_OBS_VALUES.has(normalizedObs);
}

function getRowDescriptionValue(row) {
  if (row['DESCRIÃ‡ÃƒO'] !== undefined) {
    return row['DESCRIÃ‡ÃƒO'];
  }
  if (row['DESCRIÃ‡ÃƒO ITEM'] !== undefined) {
    return row['DESCRIÃ‡ÃƒO ITEM'];
  }
  return '';
}

function getRowObsValue(row) {
  if (row.OBS !== undefined) {
    return row.OBS;
  }
  if (row.STATUS !== undefined) {
    return row.STATUS;
  }
  return '';
}

function getReqKeyForCurrentList() {
  if (state.loadedColumns.includes('N REQ')) {
    return 'N REQ';
  }
  if (state.loadedColumns.includes('REQ')) {
    return 'REQ';
  }
  return null;
}

function buildDuplicatedReqSet(rows) {
  const reqKey = getReqKeyForCurrentList();
  if (!reqKey) {
    return new Set();
  }

  const reqCounter = new Map();
  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    const normalizedReq = normalizeForFilter(row[reqKey]);
    if (!normalizedReq) {
      return;
    }
    reqCounter.set(normalizedReq, (reqCounter.get(normalizedReq) || 0) + 1);
  });

  return new Set(
    Array.from(reqCounter.entries())
      .filter(([, count]) => count > 1)
      .map(([req]) => req)
  );
}

function isCriticalRowForCurrentList(row, duplicatedReqs = null) {
  const normalizedDescricao = normalizeForFilter(getRowDescriptionValue(row));
  const normalizedObs = normalizeForFilter(getRowObsValue(row));
  const reqKey = getReqKeyForCurrentList();
  const normalizedReq = reqKey ? normalizeForFilter(row[reqKey]) : '';
  const hasCriticalInDescricao = normalizedDescricao.includes('C/ COR') || normalizedDescricao.includes('RANHURA');
  const hasCriticalInObs = normalizedObs.includes('CORTE') || normalizedObs.includes('RANHURA');
  const hasDuplicatedReq = duplicatedReqs instanceof Set && normalizedReq !== '' && duplicatedReqs.has(normalizedReq);
  return hasCriticalInDescricao || hasCriticalInObs || hasDuplicatedReq;
}

function isRowMissingForCurrentList(row) {
  if (isProjeto()) {
    return !isProjectReleasedObs(row.OBS);
  }

  if (isAcabamento()) {
    return !isAcabamentoReadyObs(row.OBS);
  }

  if (isMachineList()) {
    return normalizeForFilter(row.OBS) !== 'PRONTO';
  }

  if (isMoldagem()) {
    const status = normalizeForFilter(row.STATUS);
    // BUCHA ESTOQUE nao deve entrar como pendente.
    return status !== MOLDAGEM_STATUS_PRENSADO && status !== MOLDAGEM_STATUS_ESTOQUE;
  }

  if (state.currentListType === 'mestra') {
    return normalizeForFilter(row.OBS) !== 'PRONTO';
  }

  return false;
}

function getVisibleColumns(columns) {
  if (!isMoldagem()) {
    return columns;
  }

  return columns.filter((col) => !MOLDAGEM_HIDDEN_COLUMNS.has(normalizeColumnName(col)));
}

function shouldCenterColumn(columnName) {
  const normalized = normalizeColumnName(columnName);
  if (isMoldagem()) {
    return MOLDAGEM_CENTER_COLUMNS.has(normalized);
  }

  return STANDARD_CENTER_COLUMNS.has(normalized);
}

function updateMoldagemFilterVisibility() {
  const container = elements.filterBuchaNotEmpty.closest('.field');
  if (!container) {
    return;
  }

  if (isMoldagem()) {
    elements.filterBuchaNotEmpty.checked = false;
    container.style.display = 'none';
    return;
  }

  container.style.display = 'grid';
}

function updateProjectMissingFilterVisibility() {
  const container = elements.filterProjectMissing.closest('.field');
  if (!container) {
    return;
  }

  container.style.display = 'flex';
}

function updateAcabamentoCriticalFilterVisibility() {
  const container = elements.filterAcabamentoCritical.closest('.field');
  if (!container) {
    return;
  }

  if (isMoldagem()) {
    elements.filterAcabamentoCritical.checked = false;
    container.style.display = 'none';
    return;
  }

  container.style.display = 'flex';
}

function updateDynamicControlsVisibility() {
  const show = isMachineList();
  const buttonContainer = elements.dynamicEffSwitch.closest('.field');
  const hourContainer = elements.dynamicHourInput.closest('.field');

  if (buttonContainer) {
    buttonContainer.style.display = show ? 'flex' : 'none';
  }
  if (hourContainer) {
    hourContainer.style.display = show ? 'grid' : 'none';
  }

  if (!show) {
    state.dynamicEfficiencyEnabled = false;
    elements.dynamicEffSwitch.checked = false;
    elements.dynamicHourInput.value = '';
  }
}

function showToast(message, isError = false) {
  if (state.toastTimer) {
    clearTimeout(state.toastTimer);
  }

  elements.toast.textContent = message;
  elements.toast.style.borderLeftColor = isError ? '#b42318' : '';
  elements.toast.classList.add('show');

  state.toastTimer = setTimeout(() => {
    elements.toast.classList.remove('show');
  }, 3000);
}

function openAdminSettingsModal() {
  if (!elements.adminSettingsModal) {
    return;
  }

  elements.adminSettingsModal.hidden = false;
  if (elements.adminSettingsError) {
    elements.adminSettingsError.textContent = '';
  }
  if (elements.adminSettingsPassword) {
    elements.adminSettingsPassword.value = '';
    elements.adminSettingsPassword.focus();
  }
}

function closeAdminSettingsModal() {
  if (!elements.adminSettingsModal) {
    return;
  }

  elements.adminSettingsModal.hidden = true;
  if (elements.adminSettingsError) {
    elements.adminSettingsError.textContent = '';
  }
  if (elements.adminSettingsPassword) {
    elements.adminSettingsPassword.value = '';
  }
}

function formatBackupDateLabel(value) {
  const iso = String(value || '').trim();
  if (!iso) {
    return '';
  }
  const dt = new Date(iso);
  if (Number.isNaN(dt.getTime())) {
    return '';
  }
  return dt.toLocaleString('pt-BR');
}

function closeBackupRestoreModal() {
  if (!elements.backupRestoreModal) {
    return;
  }
  elements.backupRestoreModal.hidden = true;
  if (elements.backupRestoreError) {
    elements.backupRestoreError.textContent = '';
  }
  [elements.backupRestorePass1, elements.backupRestorePass2, elements.backupRestorePass3]
    .filter(Boolean)
    .forEach((el) => { el.value = ''; });
}

async function openBackupRestoreModal() {
  if (!elements.backupRestoreModal || !elements.backupRestoreSelect) {
    return;
  }
  elements.backupRestoreError.textContent = '';
  elements.backupRestoreModal.hidden = false;
  [elements.backupRestorePass1, elements.backupRestorePass2, elements.backupRestorePass3]
    .filter(Boolean)
    .forEach((el) => {
      el.disabled = false;
      el.readOnly = false;
      el.value = '';
    });
  elements.backupRestoreSelect.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Carregando backups...';
  elements.backupRestoreSelect.appendChild(placeholder);

  try {
    const result = await window.api.backupList();
    const rows = Array.isArray(result?.rows) ? result.rows : [];
    elements.backupRestoreSelect.innerHTML = '';
    const first = document.createElement('option');
    first.value = '';
    first.textContent = rows.length ? 'Selecione um backup' : 'Nenhum backup disponível';
    elements.backupRestoreSelect.appendChild(first);
    rows.forEach((row) => {
      const opt = document.createElement('option');
      opt.value = String(row.folderName || '').trim();
      const dateLabel = formatBackupDateLabel(row.createdAt);
      const filesLabel = Number(row.copiedCount || 0);
      opt.textContent = `${row.folderName}${dateLabel ? ` | ${dateLabel}` : ''} | ${filesLabel} arquivo(s)`;
      elements.backupRestoreSelect.appendChild(opt);
    });
    if (rows.length) {
      elements.backupRestoreSelect.value = String(rows[0].folderName || '').trim();
    }
    if (elements.backupRestorePass1) {
      elements.backupRestorePass1.focus();
    }
  } catch (error) {
    elements.backupRestoreError.textContent = error.message || 'Erro ao listar backups.';
  }
}

function setLoginsSectionVisible(visible) {
  const show = !!visible;
  state.isLoginsSectionUnlocked = show;
  if (elements.loginsAdminSection) {
    elements.loginsAdminSection.hidden = !show;
  }
  if (elements.loginsLockedNote) {
    elements.loginsLockedNote.hidden = show;
  }
  if (elements.btnUnlockLogins) {
    elements.btnUnlockLogins.textContent = show ? 'Logins Desbloqueados' : 'Desbloquear Logins';
    elements.btnUnlockLogins.disabled = show;
  }
}

function requireAdminSettings(action) {
  if (state.adminSettingsPassword) {
    return action();
  }

  state.adminPendingAction = action;
  openAdminSettingsModal();
  return null;
}

function setUsersHeaders() {
  if (!elements.usersTableHead) {
    return;
  }
  const dynamicCols = [];
  getManagedAreas().forEach((area) => {
    dynamicCols.push(area.label);
    if (area.key === 'rh') {
      dynamicCols.push('RH Editar');
    }
    if (area.key === 'fiscal') {
      FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
        dynamicCols.push(perm.title.replace('Pode ', ''));
      });
    }
  });
  setTableHeaders(elements.usersTableHead, ['Usuário', ...dynamicCols, 'Ações']);
}

function openUserModal() {
  if (!elements.userModal) {
    return;
  }
  elements.userModal.hidden = false;
  elements.userModalError.textContent = '';
  elements.userUsername.value = '';
  elements.userPassword.value = '';
  renderUserPermissionsEditor({});
  elements.userUsername.focus();
}

function closeUserModal() {
  if (!elements.userModal) {
    return;
  }
  elements.userModal.hidden = true;
  elements.userModalError.textContent = '';
}

function openUserPassModal(username) {
  if (!elements.userPassModal) {
    return;
  }
  state.userPassTarget = username;
  elements.userPassModal.hidden = false;
  elements.userPassError.textContent = '';
  elements.userPassInput.value = '';
  elements.userPassSubtitle.textContent = `Usuário: ${username}`;
  elements.userPassInput.focus();
}

function closeUserPassModal() {
  if (!elements.userPassModal) {
    return;
  }
  elements.userPassModal.hidden = true;
  elements.userPassError.textContent = '';
  elements.userPassInput.value = '';
  state.userPassTarget = null;
}

async function loadUsers() {
  const fiscalRoot = String(state.settings?.fiscalRoot || '').trim();
  if (!fiscalRoot) {
    throw new Error('Configure o Caminho Fiscal (R:) para gerenciar logins.');
  }
  const result = await window.api.listUsers();
  state.users = (result && result.users) || [];
  if (state.authUser && state.authUser.username) {
    const current = state.users.find((item) => String(item.username || '').trim() === String(state.authUser.username || '').trim());
    if (current) {
      state.authUser.permissions = current.permissions || {};
    }
  }
}

function renderUsersTable() {
  if (!elements.usersTableBody || !elements.usersEmpty || !elements.usersTableHead) {
    return;
  }

  setUsersHeaders();
  elements.usersTableBody.innerHTML = '';

  const users = state.users || [];
  if (!users.length) {
    elements.usersEmpty.style.display = 'block';
    return;
  }

  elements.usersEmpty.style.display = 'none';
  const managedAreas = getManagedAreas();
  users.forEach((user) => {
    const tr = document.createElement('tr');

    const userTd = document.createElement('td');
    userTd.textContent = user.username;
    tr.appendChild(userTd);

    const renderPermissionToggleCell = (permKey, checkedValue) => {
      const areaTd = document.createElement('td');
      areaTd.className = 'users-perm-cell';

      const toggleLabel = document.createElement('label');
      toggleLabel.className = 'users-perm-toggle';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'users-table-checkbox';
      checkbox.checked = !!checkedValue;
      checkbox.addEventListener('change', async () => {
        const nextValue = checkbox.checked;
        checkbox.checked = !nextValue;
        await requireAdminSettings(async () => {
          try {
            const nextPermissions = { ...(user.permissions || {}) };
            nextPermissions[permKey] = nextValue;

            // Mantém consistência: não faz sentido editar RH sem poder abrir RH.
            if (permKey === 'rh_edit' && nextValue) {
              nextPermissions.rh = true;
            }
            // Se tirar RH, tira também a edição.
            if (permKey === 'rh' && !nextValue) {
              nextPermissions.rh_edit = false;
            }
            if (permKey === 'fiscal' && !nextValue) {
              FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
                nextPermissions[perm.key] = false;
              });
            }
            if (permKey.startsWith('compras_') && nextValue) {
              nextPermissions.fiscal = true;
            }

            await window.api.updateUser({
              username: user.username,
              permissions: nextPermissions,
              adminPassword: state.adminSettingsPassword
            });
            await loadUsers();
            renderUsersTable();
            applyPermissionsToTabs();
            showToast('Permissão atualizada.');
          } catch (error) {
            showToast(error.message || 'Erro ao atualizar permissão.', true);
          }
        });
      });

      const slider = document.createElement('span');
      slider.className = 'users-perm-toggle-slider';

      toggleLabel.appendChild(checkbox);
      toggleLabel.appendChild(slider);
      areaTd.appendChild(toggleLabel);
      return areaTd;
    };

    managedAreas.forEach((area) => {
      tr.appendChild(renderPermissionToggleCell(area.key, getPermissionValue(user.permissions || {}, area.key)));
      if (area.key === 'rh') {
        tr.appendChild(renderPermissionToggleCell('rh_edit', getPermissionValue(user.permissions || {}, 'rh_edit')));
      }
      if (area.key === 'fiscal') {
        FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
          tr.appendChild(renderPermissionToggleCell(
            perm.key,
            getPermissionValueWithFallback(user.permissions || {}, perm.key, getPermissionValue(user.permissions || {}, 'fiscal'))
          ));
        });
      }
    });

    const actionsTd = document.createElement('td');
    const resetBtn = createPcpActionButton('Senha', () => {
      requireAdminSettings(() => openUserPassModal(user.username));
    });
    const deleteBtn = createPcpActionButton('Excluir', async () => {
      await requireAdminSettings(async () => {
        try {
          await window.api.deleteUser({ username: user.username, adminPassword: state.adminSettingsPassword });
          await loadUsers();
          renderUsersTable();
          applyPermissionsToTabs();
          showToast('Usuário excluído.');
        } catch (error) {
          showToast(error.message || 'Erro ao excluir usuário.', true);
        }
      });
    });
    actionsTd.appendChild(resetBtn);
    actionsTd.appendChild(document.createTextNode(' | '));
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(actionsTd);
    elements.usersTableBody.appendChild(tr);
  });
}

function switchTab(next) {
  const tabMap = {
    menu: { tab: elements.tabMenu, panel: elements.panelMenu },
    main: { tab: elements.tabMain, panel: elements.panelMain },
    pcp: { tab: elements.tabPcp, panel: elements.panelPcp },
    fiscal: { tab: elements.tabFiscal, panel: elements.panelFiscal },
    rh: { tab: elements.tabRh, panel: elements.panelRh },
    settings: { tab: elements.tabSettings, panel: elements.panelSettings },
    log: { tab: elements.tabLog, panel: elements.panelLog }
  };

  Object.entries(tabMap).forEach(([key, item]) => {
    if (!item.tab || !item.panel) {
      return;
    }
    const isActive = key === next;
    item.tab.classList.toggle('active', isActive);
    item.panel.classList.toggle('active', isActive);
  });
}

function getAuthPermissions() {
  return (state.authUser && state.authUser.permissions) || {};
}

function canAccessArea(areaKey) {
  return getPermissionValue(getAuthPermissions(), areaKey);
}

function canAccessPcp() {
  return canAccessArea('pcp');
}

function canAccessFiscal() {
  return canAccessArea('fiscal');
}

function applyPermissionsToTabs() {
  getManagedAreas().forEach((area) => {
    const tabEl = document.querySelector(`.tabs .tab[data-tab="${area.key}"]`);
    if (!tabEl) {
      return;
    }
    const allowed = canAccessArea(area.key);
    tabEl.disabled = !allowed;
    tabEl.title = allowed ? '' : `Usuário sem permissão para ${area.label}.`;
  });

  const menuMap = [
    { el: elements.menuOpenMain, key: 'main', label: 'Listagem' },
    { el: elements.menuOpenPcp, key: 'pcp', label: 'PCP' },
    { el: elements.menuOpenFiscal, key: 'fiscal', label: 'Fiscal' },
    { el: elements.menuOpenRh, key: 'rh', label: 'RH' },
    { el: elements.menuOpenSettings, key: 'settings', label: 'Configuração' },
    { el: elements.menuOpenLog, key: 'log', label: 'LOG' }
  ];
  menuMap.forEach(({ el, key, label }) => {
    if (!el) {
      return;
    }
    const allowed = canAccessArea(key);
    el.disabled = !allowed;
    el.title = allowed ? '' : `Usuário sem permissão para ${label}.`;
  });
}

function openAppLoginModal() {
  if (!elements.appLoginModal) {
    return;
  }

  elements.appLoginModal.hidden = false;
  if (elements.appLoginError) {
    elements.appLoginError.textContent = '';
  }
  if (elements.appLoginUsername) {
    const last = String(state.settings?.lastLoginUsername || '').trim();
    elements.appLoginUsername.value = last;
  }
  if (elements.appLoginPassword) {
    elements.appLoginPassword.value = '';
  }
  if (elements.appLoginUsername && elements.appLoginUsername.value) {
    elements.appLoginPassword?.focus?.();
  } else {
    elements.appLoginUsername?.focus?.();
  }
}

function closeAppLoginModal() {
  if (!elements.appLoginModal) {
    return;
  }

  elements.appLoginModal.hidden = true;
}

function getNowDatetimeLocal() {
  const now = new Date();
  const offsetMs = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offsetMs).toISOString().slice(0, 16);
}

function openFiscalNfModal() {
  if (!elements.fiscalNfModal) {
    return;
  }

  state.fiscalPreviewItems = [];
  elements.fiscalNfModal.hidden = false;
  if (elements.fiscalNfError) {
    elements.fiscalNfError.textContent = '';
  }
  if (elements.fiscalNfPreview) {
    elements.fiscalNfPreview.hidden = true;
  }
  if (elements.fiscalNfPreviewTitle) {
    elements.fiscalNfPreviewTitle.textContent = '';
  }
  if (elements.fiscalNfPreviewHead) {
    elements.fiscalNfPreviewHead.innerHTML = '';
  }
  if (elements.fiscalNfPreviewBody) {
    elements.fiscalNfPreviewBody.innerHTML = '';
  }
  if (elements.fiscalNfConfirm) {
    elements.fiscalNfConfirm.disabled = true;
  }
  if (elements.fiscalNfDatetime) {
    elements.fiscalNfDatetime.value = getNowDatetimeLocal();
  }
  refreshFiscalClientSuggestions();
  hideFiscalClientSuggestions();
  elements.fiscalNfNumber?.focus?.();
}

function closeFiscalNfModal() {
  if (!elements.fiscalNfModal) {
    return;
  }

  elements.fiscalNfModal.hidden = true;
  if (elements.fiscalNfNumber) {
    elements.fiscalNfNumber.value = '';
  }
  if (elements.fiscalNfClient) {
    elements.fiscalNfClient.value = '';
  }
  if (elements.fiscalNfIdentifiers) {
    elements.fiscalNfIdentifiers.value = '';
  }
  if (elements.fiscalNfError) {
    elements.fiscalNfError.textContent = '';
  }
  if (elements.fiscalNfPreview) {
    elements.fiscalNfPreview.hidden = true;
  }
  if (elements.fiscalNfConfirm) {
    elements.fiscalNfConfirm.disabled = true;
  }
  state.fiscalPreviewItems = [];
}

function parseFiscalIdentifiers(raw) {
  const tokens = String(raw || '')
    .split(/[\n,;]+/g)
    .map((item) => item.trim())
    .filter(Boolean)
    .flatMap((item) => item.split(/\s+/g).map((piece) => piece.trim()).filter(Boolean));

  const unique = [];
  const seen = new Set();
  tokens.forEach((token) => {
    const normalized = token.replace(/^#/, '');
    if (!seen.has(normalized)) {
      seen.add(normalized);
      unique.push(normalized);
    }
  });
  return unique;
}

function setTableHeaders(headEl, columns) {
  if (!headEl) {
    return;
  }
  headEl.innerHTML = '';
  const tr = document.createElement('tr');
  columns.forEach((name) => {
    const th = document.createElement('th');
    th.textContent = name;
    tr.appendChild(th);
  });
  headEl.appendChild(tr);
}

function renderFiscalPreview(items) {
  if (!elements.fiscalNfPreview || !elements.fiscalNfPreviewBody || !elements.fiscalNfPreviewHead) {
    return;
  }

  setTableHeaders(elements.fiscalNfPreviewHead, ['QNTD', 'PRODUTO', 'PEDIDO', 'REQ', 'DATA ENTRADA']);
  elements.fiscalNfPreviewBody.innerHTML = '';

  items.forEach((item) => {
    const tr = document.createElement('tr');
    [item.qtd, item.produto, item.pedido, item.req, item.dataEntrada].forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    elements.fiscalNfPreviewBody.appendChild(tr);
  });

  elements.fiscalNfPreview.hidden = false;
  if (elements.fiscalNfPreviewTitle) {
    elements.fiscalNfPreviewTitle.textContent = `${formatInteger(items.length)} item(ns) encontrados.`;
  }
}

function renderFiscalTable(items) {
  if (!elements.fiscalTableBody || !elements.fiscalTableHead || !elements.fiscalEmptyState) {
    return;
  }

  elements.fiscalTableBody.innerHTML = '';
  if (!items || !items.length) {
    elements.fiscalEmptyState.style.display = 'block';
    return;
  }

  elements.fiscalEmptyState.style.display = 'none';
  setTableHeaders(elements.fiscalTableHead, ['QNTD', 'PRODUTO', 'PEDIDO', 'CLIENTE', 'NF', 'STATUS', 'DATA FATURADA', 'HORÁRIO']);
  items.forEach((item) => {
    const tr = document.createElement('tr');
    const values = [
      item.qtd,
      item.produto,
      item.pedido,
      item.cliente,
      item.nf,
      item.status,
      item.dataFaturada,
      item.horario
    ];
    values.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    elements.fiscalTableBody.appendChild(tr);
  });
}

function openFiscalEditModal() {
  if (!elements.fiscalEditModal) {
    return;
  }

  const nf = state.fiscalSelectionNf;
  if (!nf) {
    return;
  }

  const nfInfo = (state.fiscalNfs || []).find((item) => String(item.nf) === String(nf)) || null;
  if (elements.fiscalEditStatus) {
    elements.fiscalEditStatus.value = (nfInfo && nfInfo.status) ? String(nfInfo.status) : 'Faturado';
  }
  if (elements.fiscalEditRastreio) {
    elements.fiscalEditRastreio.value = (nfInfo && nfInfo.rastreio) ? String(nfInfo.rastreio) : '';
  }
  if (elements.fiscalEditDespache) {
    elements.fiscalEditDespache.value = '';
  }
  if (elements.fiscalEditError) {
    elements.fiscalEditError.textContent = '';
  }

  elements.fiscalEditModal.hidden = false;
}

function closeFiscalEditModal() {
  if (!elements.fiscalEditModal) {
    return;
  }
  elements.fiscalEditModal.hidden = true;
  if (elements.fiscalEditError) {
    elements.fiscalEditError.textContent = '';
  }
}

function openFiscalConfirmModal() {
  if (!elements.fiscalConfirmModal) {
    return;
  }

  elements.fiscalConfirmModal.hidden = false;
  if (elements.fiscalConfirmError) {
    elements.fiscalConfirmError.textContent = '';
  }
  if (elements.fiscalConfirmUserLabel) {
    elements.fiscalConfirmUserLabel.textContent = `Usuário: ${String(state.authUser?.username || state.fiscalUser || '').trim()}`;
  }
  elements.fiscalConfirmOk?.focus?.();
}

function closeFiscalConfirmModal() {
  if (!elements.fiscalConfirmModal) {
    return;
  }
  elements.fiscalConfirmModal.hidden = true;
  if (elements.fiscalConfirmError) {
    elements.fiscalConfirmError.textContent = '';
  }
}

function openFiscalSearchModal() {
  if (!elements.fiscalSearchModal) {
    return;
  }

  elements.fiscalSearchModal.hidden = false;
  if (elements.fiscalSearchError) {
    elements.fiscalSearchError.textContent = '';
  }
  if (elements.fiscalSearchNf) {
    elements.fiscalSearchNf.value = '';
    elements.fiscalSearchNf.focus();
  }
}

function closeFiscalSearchModal() {
  if (!elements.fiscalSearchModal) {
    return;
  }

  elements.fiscalSearchModal.hidden = true;
  if (elements.fiscalSearchError) {
    elements.fiscalSearchError.textContent = '';
  }
  if (elements.fiscalSearchNf) {
    elements.fiscalSearchNf.value = '';
  }
}

function openFiscalDeleteModal() {
  if (!elements.fiscalDeleteModal) {
    return;
  }

  if (!state.fiscalSelectionNf) {
    return;
  }

  state.fiscalDeleteReason = '';
  elements.fiscalDeleteModal.hidden = false;
  if (elements.fiscalDeleteError) {
    elements.fiscalDeleteError.textContent = '';
  }
  if (elements.fiscalDeleteReason) {
    elements.fiscalDeleteReason.value = '';
    elements.fiscalDeleteReason.focus();
  }
}

function closeFiscalDeleteModal() {
  if (!elements.fiscalDeleteModal) {
    return;
  }

  elements.fiscalDeleteModal.hidden = true;
  if (elements.fiscalDeleteError) {
    elements.fiscalDeleteError.textContent = '';
  }
  if (elements.fiscalDeleteReason) {
    elements.fiscalDeleteReason.value = '';
  }
}

function reopenFiscalEditModal() {
  if (!elements.fiscalEditModal) {
    return;
  }
  elements.fiscalEditModal.hidden = false;
}

function reopenFiscalDeleteModal() {
  if (!elements.fiscalDeleteModal) {
    return;
  }

  elements.fiscalDeleteModal.hidden = false;
  if (elements.fiscalDeleteError) {
    elements.fiscalDeleteError.textContent = '';
  }
  if (elements.fiscalDeleteReason && !String(elements.fiscalDeleteReason.value || '').trim()) {
    elements.fiscalDeleteReason.value = state.fiscalDeleteReason || '';
  }
}

function renderFiscalBreadcrumb() {
  if (!elements.fiscalBreadcrumb) {
    return;
  }

  elements.fiscalBreadcrumb.innerHTML = '';
  if (state.fiscalView !== 'nf' || !state.fiscalSelectionNf) {
    elements.fiscalBreadcrumb.style.display = 'none';
    return;
  }

  elements.fiscalBreadcrumb.style.display = 'flex';
  const steps = [{ label: 'NFs', level: 'nfs' }, { label: `NF ${state.fiscalSelectionNf}`, level: 'nf' }];
  steps.forEach((step, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'pcp-crumb';
    button.textContent = step.label;
    button.addEventListener('click', () => {
      if (step.level === 'nfs') {
        state.fiscalSelectionNf = null;
        state.fiscalNfItems = [];
        renderFiscalScreen();
      }
    });
    elements.fiscalBreadcrumb.appendChild(button);

    if (index < steps.length - 1) {
      const divider = document.createElement('span');
      divider.className = 'pcp-crumb-divider';
      divider.textContent = '/';
      elements.fiscalBreadcrumb.appendChild(divider);
    }
  });
}

function renderFiscalNfList() {
  if (!elements.fiscalTableBody || !elements.fiscalTableHead || !elements.fiscalEmptyState) {
    return;
  }

  const nfs = state.fiscalNfs || [];
  elements.fiscalTableBody.innerHTML = '';
  setTableHeaders(elements.fiscalTableHead, ['NF', 'CLIENTE', 'PEDIDOS', 'REQUISIÇÕES', 'STATUS', 'DATA FATURADA', 'RASTREIO', 'DATA DESPACHE']);

  if (!nfs.length) {
    elements.fiscalEmptyState.style.display = 'block';
    elements.fiscalEmptyState.textContent = 'Nenhuma NF encontrada no banco de pedidos.';
    return;
  }

  elements.fiscalEmptyState.style.display = 'none';
  nfs.forEach((item) => {
    const tr = document.createElement('tr');

    const nfTd = document.createElement('td');
    nfTd.appendChild(createPcpActionButton(String(item.nf || ''), async () => {
      await openFiscalNfDetails(item.nf);
    }));

    const values = [
      String(item.cliente || ''),
      String(item.pedidosLabel || ''),
      String(item.reqsLabel || ''),
      String(item.status || ''),
      String(item.dataFaturada || ''),
      String(item.rastreio || ''),
      String(item.dataDespache || '')
    ];

    tr.appendChild(nfTd);
    values.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });

    elements.fiscalTableBody.appendChild(tr);
  });
}

function isFiscalPendingNf(item) {
  const status = String(item && item.status ? item.status : '').trim().toUpperCase();
  const rastreio = String(item && item.rastreio ? item.rastreio : '').trim();
  return !rastreio || status === 'FATURADO';
}

function updateFiscalPendingButtonCount() {
  if (!elements.fiscalBtnPendencias) {
    return;
  }
  const pendingCount = (state.fiscalNfs || []).filter(isFiscalPendingNf).length;
  elements.fiscalBtnPendencias.textContent = `Pendências (${formatInteger(pendingCount)})`;
}

function renderFiscalPendingNfList() {
  if (!elements.fiscalTableBody || !elements.fiscalTableHead || !elements.fiscalEmptyState) {
    return;
  }

  const pending = (state.fiscalNfs || []).filter(isFiscalPendingNf);
  elements.fiscalTableBody.innerHTML = '';
  setTableHeaders(elements.fiscalTableHead, ['NF', 'CLIENTE', 'STATUS', 'RASTREIO', 'DATA FATURADA', 'PEDIDOS']);

  if (!pending.length) {
    elements.fiscalEmptyState.style.display = 'block';
    elements.fiscalEmptyState.textContent = 'Nenhuma pendência encontrada.';
    return;
  }

  elements.fiscalEmptyState.style.display = 'none';
  pending.forEach((item) => {
    const tr = document.createElement('tr');

    const nfTd = document.createElement('td');
    nfTd.appendChild(createPcpActionButton(String(item.nf || ''), async () => {
      await openFiscalNfDetails(item.nf);
    }));
    tr.appendChild(nfTd);

    const values = [
      String(item.cliente || ''),
      String(item.status || ''),
      String(item.rastreio || ''),
      String(item.dataFaturada || ''),
      String(item.pedidosLabel || '')
    ];
    values.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });
    elements.fiscalTableBody.appendChild(tr);
  });
}

function renderFiscalNfItems() {
  if (!elements.fiscalTableBody || !elements.fiscalTableHead || !elements.fiscalEmptyState) {
    return;
  }

  const items = state.fiscalNfItems || [];
  elements.fiscalTableBody.innerHTML = '';
  setTableHeaders(elements.fiscalTableHead, ['QNTD', 'PRODUTO', 'PEDIDO', 'CLIENTE', 'STATUS', 'DATA ENTRADA', 'DATA DESPACHE', 'NF', 'RASTREIO', 'BAIXADO POR', 'HORÁRIO']);

  if (!items.length) {
    elements.fiscalEmptyState.style.display = 'block';
    elements.fiscalEmptyState.textContent = 'Nenhum item encontrado para esta NF.';
    return;
  }

  elements.fiscalEmptyState.style.display = 'none';
  items.forEach((item) => {
    const tr = document.createElement('tr');
    const values = [
      item.qtd,
      item.produto,
      item.pedido,
      item.cliente,
      item.status,
      item.dataEntrada,
      item.dataDespache,
      item.nf,
      item.rastreio,
      item.baixadoPor,
      item.horario
    ];

    values.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    elements.fiscalTableBody.appendChild(tr);
  });
}

function renderFiscalHistory() {
  if (!elements.fiscalTableBody || !elements.fiscalTableHead || !elements.fiscalEmptyState) {
    return;
  }

  const rows = state.fiscalHistoryRows || [];
  elements.fiscalTableBody.innerHTML = '';
  setTableHeaders(elements.fiscalTableHead, ['NF', 'APAGADO POR', 'DATA', 'HORA', 'MOTIVO']);

  if (!rows.length) {
    elements.fiscalEmptyState.style.display = 'block';
    elements.fiscalEmptyState.textContent = 'Nenhum historico de NF apagada encontrado.';
    return;
  }

  elements.fiscalEmptyState.style.display = 'none';
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    const values = [row.nf, row.usuario, row.data, row.hora, row.motivo];
    values.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    elements.fiscalTableBody.appendChild(tr);
  });
}

function renderFiscalScreen() {
  if (!elements.panelFiscal) {
    return;
  }

  const isMenu = state.fiscalView === 'menu';
  const isCompras = state.fiscalView === 'compras';
  if (elements.fiscalMenuView) {
    elements.fiscalMenuView.hidden = !isMenu;
  }
  if (elements.fiscalNfsView) {
    elements.fiscalNfsView.hidden = isMenu || isCompras;
  }
  if (elements.fiscalComprasView) {
    elements.fiscalComprasView.hidden = !isCompras;
  }
  if (isMenu || isCompras) {
    return;
  }

  const hasSelection = !!state.fiscalSelectionNf && state.fiscalView === 'nf';
  const isHistory = state.fiscalView === 'history';
  const isPending = state.fiscalView === 'pending';
  if (elements.fiscalBtnVoltar) {
    elements.fiscalBtnVoltar.hidden = !hasSelection && !isHistory && !isPending;
  }
  if (elements.fiscalBtnEditarNf) {
    elements.fiscalBtnEditarNf.hidden = !hasSelection;
  }
  if (elements.fiscalBtnApagarNf) {
    elements.fiscalBtnApagarNf.hidden = !hasSelection;
  }

  renderFiscalBreadcrumb();
  if (isHistory) {
    if (elements.fiscalSubtitle) {
      elements.fiscalSubtitle.textContent = 'Historico de NFs apagadas.';
    }
    renderFiscalHistory();
    return;
  }

  if (isPending) {
    const pendingCount = (state.fiscalNfs || []).filter(isFiscalPendingNf).length;
    if (elements.fiscalSubtitle) {
      elements.fiscalSubtitle.textContent = `Pendências fiscais (${formatInteger(pendingCount)}) - sem rastreio ou com status Faturado.`;
    }
    renderFiscalPendingNfList();
    return;
  }

  if (!hasSelection) {
    if (elements.fiscalSubtitle) {
      elements.fiscalSubtitle.textContent = 'NFs emitidas no banco de pedidos.';
    }
    renderFiscalNfList();
    return;
  }

  if (elements.fiscalSubtitle) {
    elements.fiscalSubtitle.textContent = `Itens da NF ${state.fiscalSelectionNf}`;
  }
  renderFiscalNfItems();
}

async function loadFiscalNfs() {
  const result = await window.api.fiscalListNfs();
  state.fiscalNfs = (result && result.nfs) || [];
  refreshFiscalClientSuggestions();
  updateFiscalPendingButtonCount();
}

function getFiscalClientNames() {
  const unique = new Map();
  (state.fiscalNfs || []).forEach((item) => {
    const client = String(item && item.cliente ? item.cliente : '').trim();
    if (!client) {
      return;
    }
    const key = client.toLocaleUpperCase('pt-BR');
    if (!unique.has(key)) {
      unique.set(key, client);
    }
  });

  return Array.from(unique.values()).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function hideFiscalClientSuggestions() {
  if (!elements.fiscalClientSuggestions) {
    return;
  }
  elements.fiscalClientSuggestions.hidden = true;
  elements.fiscalClientSuggestions.innerHTML = '';
}

function refreshFiscalClientSuggestions() {
  if (!elements.fiscalNfClient) {
    return;
  }
  renderFiscalClientSuggestions(elements.fiscalNfClient.value || '');
}

function renderFiscalClientSuggestions(rawQuery) {
  if (!elements.fiscalClientSuggestions) {
    return;
  }

  const query = String(rawQuery || '').trim();
  if (!query) {
    hideFiscalClientSuggestions();
    return;
  }

  const normalizedQuery = normalizeForFilter(query);
  const matches = getFiscalClientNames()
    .filter((name) => normalizeForFilter(name).includes(normalizedQuery))
    .slice(0, 8);

  if (!matches.length) {
    hideFiscalClientSuggestions();
    return;
  }

  elements.fiscalClientSuggestions.innerHTML = '';
  matches.forEach((name) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'fiscal-client-option';
    button.textContent = name;
    button.addEventListener('click', () => {
      if (elements.fiscalNfClient) {
        elements.fiscalNfClient.value = name;
      }
      hideFiscalClientSuggestions();
    });
    elements.fiscalClientSuggestions.appendChild(button);
  });
  elements.fiscalClientSuggestions.hidden = false;
}

async function loadFiscalHistory() {
  const result = await window.api.fiscalHistory();
  state.fiscalHistoryRows = (result && result.rows) || [];
}

async function openFiscalNfDetails(nf) {
  const result = await window.api.fiscalGetNf({ nf });
  state.fiscalSelectionNf = String(nf || '').trim() || null;
  state.fiscalNfItems = (result && result.items) || [];
  state.fiscalView = 'nf';
  renderFiscalScreen();
}

async function handleFiscalSearchNf() {
  try {
    const nf = String(elements.fiscalSearchNf?.value || '').trim();
    if (!nf) {
      if (elements.fiscalSearchError) {
        elements.fiscalSearchError.textContent = 'Informe a NF para buscar.';
      }
      return;
    }

    if (elements.fiscalSearchError) {
      elements.fiscalSearchError.textContent = '';
    }

    const direct = await window.api.fiscalGetNf({ nf });
    const directItems = (direct && direct.items) || [];
    if (directItems.length) {
      state.fiscalSelectionNf = String(nf || '').trim() || null;
      state.fiscalNfItems = directItems;
      state.fiscalView = 'nf';
      renderFiscalScreen();
      closeFiscalSearchModal();
      return;
    }

    const found = await window.api.fiscalFindNf({ identifier: nf });
    const nfs = (found && found.nfs) || [];
    if (!nfs.length) {
      if (elements.fiscalSearchError) {
        elements.fiscalSearchError.textContent = 'Nenhuma NF encontrada para este pedido/requisição.';
      }
      return;
    }

    if (nfs.length > 1) {
      showToast(`Mais de 1 NF encontrada: ${nfs.slice(0, 6).join(' | ')}${nfs.length > 6 ? ' | ...' : ''}`);
    }

    await openFiscalNfDetails(nfs[0]);
    closeFiscalSearchModal();
  } catch (error) {
    if (elements.fiscalSearchError) {
      elements.fiscalSearchError.textContent = error.message || 'Erro ao pesquisar NF.';
    }
  }
}

function setPcpView(view) {
  const allowed = new Set(['menu', 'acomp', 'eff', 'dashboard']);
  state.pcpView = allowed.has(view) ? view : 'menu';
  if (elements.pcpMenuView) {
    elements.pcpMenuView.hidden = state.pcpView !== 'menu';
  }
  if (elements.pcpAcompView) {
    elements.pcpAcompView.hidden = state.pcpView !== 'acomp';
  }
  if (elements.pcpEffView) {
    elements.pcpEffView.hidden = state.pcpView !== 'eff';
  }
  if (elements.pcpDashboardView) {
    elements.pcpDashboardView.hidden = state.pcpView !== 'dashboard';
  }
}

function resetPcpSelection() {
  state.pcpSelection = {
    freight: null,
    client: null,
    order: null
  };
}

function resetPcpSummary() {
  state.pcpSummary = null;
  resetPcpSelection();
  renderPcpTable();
}

function setPcpHeaders(columns) {
  elements.pcpTableHead.innerHTML = '';
  const tr = document.createElement('tr');
  columns.forEach((name) => {
    const th = document.createElement('th');
    th.textContent = name;
    tr.appendChild(th);
  });
  elements.pcpTableHead.appendChild(tr);
}

function getPcpCurrentFreight() {
  if (!state.pcpSummary || !state.pcpSelection.freight) {
    return null;
  }
  return state.pcpSummary.freights.find((item) => item.id === state.pcpSelection.freight) || null;
}

function getPcpCurrentClient() {
  const freight = getPcpCurrentFreight();
  if (!freight || !state.pcpSelection.client) {
    return null;
  }
  return freight.clients.find((item) => item.id === state.pcpSelection.client) || null;
}

function getPcpCurrentOrder() {
  const client = getPcpCurrentClient();
  if (!client || !state.pcpSelection.order) {
    return null;
  }
  return client.orders.find((item) => item.id === state.pcpSelection.order) || null;
}

function renderPcpBreadcrumb() {
  elements.pcpBreadcrumb.innerHTML = '';
  if (!state.pcpSummary) {
    elements.pcpBreadcrumb.style.display = 'none';
    return;
  }

  elements.pcpBreadcrumb.style.display = 'flex';
  const steps = [{ label: 'Fretes', level: 'freight' }];
  const freight = getPcpCurrentFreight();
  const client = getPcpCurrentClient();
  const order = getPcpCurrentOrder();

  if (freight) {
    steps.push({ label: freight.name, level: 'client' });
  }
  if (client) {
    steps.push({ label: client.name, level: 'order' });
  }
  if (order) {
    steps.push({ label: order.orderLabel, level: 'item' });
  }

  steps.forEach((step, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'pcp-crumb';
    button.textContent = step.label;
    button.addEventListener('click', () => {
      if (step.level === 'freight') {
        resetPcpSelection();
      } else if (step.level === 'client') {
        state.pcpSelection.client = null;
        state.pcpSelection.order = null;
      } else if (step.level === 'order') {
        state.pcpSelection.order = null;
      }
      renderPcpTable();
    });
    elements.pcpBreadcrumb.appendChild(button);

    if (index < steps.length - 1) {
      const divider = document.createElement('span');
      divider.className = 'pcp-crumb-divider';
      divider.textContent = '/';
      elements.pcpBreadcrumb.appendChild(divider);
    }
  });
}

function createPcpActionButton(text, onClick) {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'pcp-link-button';
  button.textContent = text;
  button.addEventListener('click', onClick);
  return button;
}

function formatPcpMetrics(total, done, pending) {
  return `${formatInteger(total)} total | ${formatInteger(done)} prontas | ${formatInteger(pending)} faltando`;
}

function formatPcpOrderCount(total) {
  return `${formatInteger(total)} pedidos`;
}

function renderPcpFreightRows() {
  const rows = state.pcpSummary ? state.pcpSummary.freights : [];
  if (!rows.length) {
    elements.pcpEmptyState.textContent = 'Nenhum item com ENTREGA encontrado para esta data/lista.';
    elements.pcpEmptyState.style.display = 'block';
    return;
  }

  setPcpHeaders(['Frete', 'Clientes (pedidos)']);
  elements.pcpEmptyState.style.display = 'none';

  rows.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.pendingOrders === 0) {
      tr.classList.add('pcp-ready-row');
    }

    const freightTd = document.createElement('td');
    freightTd.appendChild(createPcpActionButton(item.name, () => {
      state.pcpSelection.freight = item.id;
      state.pcpSelection.client = null;
      state.pcpSelection.order = null;
      renderPcpTable();
    }));

    const clientsTd = document.createElement('td');
    if (!item.clients.length) {
      clientsTd.textContent = '-';
    } else {
      const list = document.createElement('div');
      list.className = 'pcp-client-list';
      item.clients.forEach((client) => {
        const row = document.createElement('div');
        row.className = 'pcp-client-list-item';

        const clientButton = createPcpActionButton(client.name, () => {
          state.pcpSelection.freight = item.id;
          state.pcpSelection.client = client.id;
          state.pcpSelection.order = null;
          renderPcpTable();
        });
        clientButton.classList.add('pcp-client-link');

        const count = document.createElement('span');
        count.className = 'pcp-client-count';
        count.textContent = formatInteger(client.totalOrders);

        row.appendChild(clientButton);
        row.appendChild(count);
        list.appendChild(row);
      });
      clientsTd.appendChild(list);
    }

    tr.appendChild(freightTd);
    tr.appendChild(clientsTd);
    elements.pcpTableBody.appendChild(tr);
  });
}

function renderPcpClientRows(freight) {
  setPcpHeaders(['Cliente', 'Pedidos']);
  elements.pcpEmptyState.style.display = 'none';

  freight.clients.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.pendingOrders === 0) {
      tr.classList.add('pcp-ready-row');
    }

    const clientTd = document.createElement('td');
    clientTd.appendChild(createPcpActionButton(item.name, () => {
      state.pcpSelection.client = item.id;
      state.pcpSelection.order = null;
      renderPcpTable();
    }));

    const totalTd = document.createElement('td');
    totalTd.textContent = formatInteger(item.totalOrders);

    tr.appendChild(clientTd);
    tr.appendChild(totalTd);
    elements.pcpTableBody.appendChild(tr);
  });
}

function renderPcpOrderRows(client) {
  setPcpHeaders(['Pedido', 'Itens', 'Prontos', 'Faltam']);
  elements.pcpEmptyState.style.display = 'none';

  client.orders.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.pendingItems === 0) {
      tr.classList.add('pcp-ready-row');
    }

    const orderTd = document.createElement('td');
    orderTd.appendChild(createPcpActionButton(item.orderLabel, () => {
      state.pcpSelection.order = item.id;
      renderPcpTable();
    }));

    const totalTd = document.createElement('td');
    totalTd.textContent = formatInteger(item.totalItems);
    const doneTd = document.createElement('td');
    doneTd.textContent = formatInteger(item.doneItems);
    const pendingTd = document.createElement('td');
    pendingTd.textContent = formatInteger(item.pendingItems);

    tr.appendChild(orderTd);
    tr.appendChild(totalTd);
    tr.appendChild(doneTd);
    tr.appendChild(pendingTd);
    elements.pcpTableBody.appendChild(tr);
  });
}

function renderPcpItemRows(order) {
  setPcpHeaders(['Linha', 'Cliente', 'Pedido', 'OP (N REQ)', 'Descricao', 'QTD', 'Entrega', 'Status']);
  elements.pcpEmptyState.style.display = 'none';

  order.items.forEach((item) => {
    const tr = document.createElement('tr');
    const statusCode = item.statusCode || 'pendente';
    tr.classList.add(`pcp-status-${statusCode}`);

    let statusLabel = 'SEM BAIXA';
    if (statusCode === 'acabamento') {
      statusLabel = 'BAIXA ACABAMENTO';
    } else if (statusCode === 'maquina') {
      statusLabel = 'BAIXA MAQUINA';
    } else if (statusCode === 'projeto') {
      statusLabel = 'BAIXA PROJETO';
    }
    const detailValue = String(item.statusValue || '').trim();
    if (detailValue && statusCode !== 'acabamento' && statusCode !== 'pendente') {
      statusLabel = `${statusLabel} - ${detailValue}`;
    }

    const values = [
      item.lineLabel,
      item.clientName,
      item.orderLabel,
      item.reqNumber,
      item.description,
      item.qtd,
      item.deliveryType,
      statusLabel
    ];

    values.forEach((value, index) => {
      const td = document.createElement('td');
      td.textContent = value || '';
      if (index === values.length - 1) {
        td.classList.add(`pcp-status-text-${statusCode}`);
      }
      tr.appendChild(td);
    });

    elements.pcpTableBody.appendChild(tr);
  });
}

function renderPcpTable() {
  if (!elements.pcpTableBody || !elements.pcpEmptyState || !elements.pcpTableHead || !elements.pcpBreadcrumb) {
    return;
  }
  if (state.pcpView !== 'acomp') {
    return;
  }

  elements.pcpTableHead.innerHTML = '';
  elements.pcpTableBody.innerHTML = '';
  renderPcpBreadcrumb();

  if (!state.pcpSummary) {
    elements.pcpEmptyState.textContent = 'Clique na aba PCP e informe a senha para carregar o resumo.';
    elements.pcpEmptyState.style.display = 'block';
    if (elements.pcpSubtitle) {
      elements.pcpSubtitle.textContent = 'Resumo de fretes com drilldown por cliente, pedido e itens.';
    }
    return;
  }

  const freight = getPcpCurrentFreight();
  const client = getPcpCurrentClient();
  const order = getPcpCurrentOrder();

  if (order) {
    if (elements.pcpSubtitle) {
      elements.pcpSubtitle.textContent = `Pedido ${order.orderLabel} | ${formatPcpMetrics(order.totalItems, order.doneItems, order.pendingItems)}`;
    }
    renderPcpItemRows(order);
    return;
  }

  if (client) {
    if (elements.pcpSubtitle) {
      elements.pcpSubtitle.textContent = `${client.name} (${freight.name}) | ${formatPcpOrderCount(client.totalOrders)}`;
    }
    renderPcpOrderRows(client);
    return;
  }

  if (freight) {
    if (elements.pcpSubtitle) {
      elements.pcpSubtitle.textContent = `${freight.name} | ${formatInteger(freight.clients.length)} clientes | ${formatPcpOrderCount(freight.totalOrders)}`;
    }
    renderPcpClientRows(freight);
    return;
  }

  if (elements.pcpSubtitle) {
    elements.pcpSubtitle.textContent = `${state.pcpSummary.listLabel} | ${state.pcpSummary.date}`;
  }
  renderPcpFreightRows();
}

async function loadPcpSummary() {
  const date = ensureDateValue();
  const listType = elements.listTypeSelect.value || 'acabamento';
  state.pcpSummary = await window.api.loadFreightSummary({ date, listType });
  resetPcpSelection();
}

async function openPcpAcompanhamento() {
  if (!state.pcpSummary) {
    try {
      await loadPcpSummary();
    } catch (error) {
      const message = String((error && error.message) || '');
      if (message.includes('coluna ENTREGA')) {
        if (elements.listTypeSelect) {
          elements.listTypeSelect.value = 'acabamento';
        }
        state.currentListType = 'acabamento';
        resetPcpSummary();
        await loadPcpSummary();
        showToast('Listagem ajustada automaticamente para ACABAMENTO no PCP.');
      } else {
        throw error;
      }
    }
  }
  setPcpView('acomp');
  renderPcpTable();
}

function resetPcpEfficiency() {
  state.pcpEfficiencySnapshot = null;
  state.pcpEfficiencySelectedMachines = [];
  if (elements.pcpEffMachineFilter) {
    elements.pcpEffMachineFilter.innerHTML = '';
  }
  if (elements.pcpEffOverall) {
    elements.pcpEffOverall.innerHTML = '';
  }
  if (elements.pcpEffGrid) {
    elements.pcpEffGrid.innerHTML = '';
  }
  if (elements.pcpEffEmpty) {
    elements.pcpEffEmpty.style.display = 'block';
    elements.pcpEffEmpty.textContent = 'Clique em Atualizar dados para gerar o painel.';
  }
}

function getPcpSelectedMachineIds(snapshotMachines) {
  const allIds = (snapshotMachines || []).map((item) => item.listType);
  if (!state.pcpEfficiencySelectedMachines.length) {
    return allIds;
  }
  const selected = new Set(state.pcpEfficiencySelectedMachines);
  const filtered = allIds.filter((id) => selected.has(id));
  return filtered.length ? filtered : allIds;
}

function renderPcpEffMachineFilter() {
  if (!elements.pcpEffMachineFilter) {
    return;
  }

  elements.pcpEffMachineFilter.innerHTML = '';
  const machines = state.pcpEfficiencySnapshot ? state.pcpEfficiencySnapshot.machines || [] : [];
  if (!machines.length) {
    elements.pcpEffMachineFilter.style.display = 'none';
    return;
  }

  elements.pcpEffMachineFilter.style.display = 'flex';
  const selectedSet = new Set(getPcpSelectedMachineIds(machines));
  machines.forEach((machine) => {
    const chip = document.createElement('label');
    chip.className = 'pcp-eff-machine-chip';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = selectedSet.has(machine.listType);
    checkbox.addEventListener('change', () => {
      const current = new Set(getPcpSelectedMachineIds(machines));
      if (checkbox.checked) {
        current.add(machine.listType);
      } else {
        current.delete(machine.listType);
      }
      if (!current.size) {
        checkbox.checked = true;
        return;
      }
      state.pcpEfficiencySelectedMachines = Array.from(current);
      renderPcpEfficiency();
    });

    const text = document.createElement('span');
    text.textContent = machine.machineName || machine.listType.toUpperCase();
    chip.appendChild(checkbox);
    chip.appendChild(text);
    elements.pcpEffMachineFilter.appendChild(chip);
  });
}

function getPcpEffHourOverride() {
  if (state.pcpEfficiencyHourMode === 'auto') {
    return null;
  }
  const parsed = Number(state.pcpEfficiencyHourMode);
  return Number.isFinite(parsed) ? clamp(parsed, 0, 23) : null;
}

function getPcpEffPlanning(dateValue) {
  const shift = getShiftInfo(dateValue);
  const hourOverride = getPcpEffHourOverride();
  const referenceMinutes = getReferenceMinutes(dateValue, hourOverride);
  const clampedWithinShift = clamp(referenceMinutes, shift.startMinutes, shift.endMinutes);
  let workedMinutes = clampedWithinShift - shift.startMinutes;
  workedMinutes -= overlapsMinutes(shift.startMinutes, clampedWithinShift, shift.lunchStart, shift.lunchEnd);
  workedMinutes = Math.max(0, workedMinutes);
  const totalShiftMinutes = (shift.endMinutes - shift.startMinutes) - (shift.lunchEnd - shift.lunchStart);
  const progress = totalShiftMinutes > 0 ? workedMinutes / totalShiftMinutes : 0;

  return {
    isWorkday: shift.isWorkday,
    plannedPercent: progress * 100,
    progress
  };
}

function getPcpEffHourLabel() {
  if (state.pcpEfficiencyHourMode !== 'auto') {
    return `${state.pcpEfficiencyHourMode}:00`;
  }

  const now = new Date();
  return `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
}

function evaluatePcpEffStatus(actualPercent, plannedPercent) {
  if (actualPercent >= plannedPercent) {
    return 'rocket';
  }
  if (actualPercent >= plannedPercent - 15) {
    return 'okay';
  }
  return 'alert';
}

function getPcpEffStatusText(status) {
  if (status === 'rocket') {
    return 'ACIMA!';
  }
  if (status === 'okay') {
    return 'NO RITMO';
  }
  return 'ABAIXO';
}

function createPcpEffMetricBlock({ title, total, done, expected, actualPercent, plannedPercent }) {
  const block = document.createElement('article');
  block.className = 'pcp-eff-metric';

  const heading = document.createElement('h4');
  heading.textContent = title;
  block.appendChild(heading);

  const info = document.createElement('p');
  info.textContent = `Total: ${formatInteger(total)} | Feitas: ${formatInteger(done)} | Esperado: ${formatInteger(expected)}`;
  block.appendChild(info);

  const percent = document.createElement('p');
  percent.className = 'pcp-eff-percent';
  percent.textContent = `Atual: ${formatPercent(actualPercent)} | Planejado: ${formatPercent(plannedPercent)}`;
  block.appendChild(percent);

  block.appendChild(createBatteryChart(actualPercent, plannedPercent));
  return block;
}

function createPcpEffStatusVisual(status, basisLabel, actualPercent, plannedPercent) {
  const wrap = document.createElement('aside');
  wrap.className = `pcp-eff-status pcp-eff-status-${status}`;

  const image = document.createElement('img');
  image.className = 'pcp-eff-status-image';
  image.src = PCP_STATUS_IMAGE_PATHS[status];
  image.alt = getPcpEffStatusText(status);
  wrap.appendChild(image);

  const title = document.createElement('p');
  title.className = 'pcp-eff-status-title';
  title.textContent = getPcpEffStatusText(status);
  wrap.appendChild(title);

  const detail = document.createElement('p');
  detail.className = 'pcp-eff-status-detail';
  detail.textContent = `${basisLabel}: ${formatPercent(actualPercent)} x ${formatPercent(plannedPercent)} planejado`;
  wrap.appendChild(detail);

  return wrap;
}

function createPcpEffColumnChart(title, doneValue, plannedValue, doneLabel, plannedLabel) {
  const safeDone = Number(doneValue || 0);
  const safePlanned = Number(plannedValue || 0);
  const maxValue = Math.max(safeDone, safePlanned, 1);
  const doneHeight = (safeDone / maxValue) * 100;
  const plannedHeight = (safePlanned / maxValue) * 100;

  const chart = document.createElement('article');
  chart.className = 'pcp-eff-col-chart';

  const heading = document.createElement('h5');
  heading.textContent = title;
  chart.appendChild(heading);

  const barsWrap = document.createElement('div');
  barsWrap.className = 'pcp-eff-col-bars';

  const doneCol = document.createElement('div');
  doneCol.className = 'pcp-eff-col-item';
  const doneBar = document.createElement('div');
  doneBar.className = 'pcp-eff-col-bar pcp-eff-col-bar-done';
  doneBar.style.height = `${clamp(doneHeight, 0, 100)}%`;
  const doneValueLabel = document.createElement('span');
  doneValueLabel.className = 'pcp-eff-col-value';
  doneValueLabel.textContent = formatInteger(safeDone);
  const doneText = document.createElement('p');
  doneText.textContent = doneLabel;
  doneCol.appendChild(doneValueLabel);
  doneCol.appendChild(doneBar);
  doneCol.appendChild(doneText);

  const plannedCol = document.createElement('div');
  plannedCol.className = 'pcp-eff-col-item';
  const plannedBar = document.createElement('div');
  plannedBar.className = 'pcp-eff-col-bar pcp-eff-col-bar-planned';
  plannedBar.style.height = `${clamp(plannedHeight, 0, 100)}%`;
  const plannedValueLabel = document.createElement('span');
  plannedValueLabel.className = 'pcp-eff-col-value';
  plannedValueLabel.textContent = formatInteger(safePlanned);
  const plannedText = document.createElement('p');
  plannedText.textContent = plannedLabel;
  plannedCol.appendChild(plannedValueLabel);
  plannedCol.appendChild(plannedBar);
  plannedCol.appendChild(plannedText);

  barsWrap.appendChild(doneCol);
  barsWrap.appendChild(plannedCol);
  chart.appendChild(barsWrap);
  return chart;
}

function getPcpMachineBasis(machine) {
  if (machine.basis) {
    return machine.basis;
  }
  return PCP_EFF_OP_BASED_TYPES.has(machine.listType) ? 'ordens' : 'pecas';
}

function renderPcpEfficiency() {
  if (!elements.pcpEffOverall || !elements.pcpEffGrid || !elements.pcpEffEmpty || !elements.pcpEffTitle) {
    return;
  }
  if (state.pcpView !== 'eff') {
    return;
  }

  elements.pcpEffOverall.innerHTML = '';
  elements.pcpEffGrid.innerHTML = '';

  const dateLabel = ensurePcpEffDateValue();
  const hourLabel = getPcpEffHourLabel();
  elements.pcpEffTitle.textContent = `Emissão ${dateLabel} às ${hourLabel}`;

  const planning = getPcpEffPlanning(dateLabel);
  if (!planning.isWorkday) {
    elements.pcpEffEmpty.style.display = 'block';
    elements.pcpEffEmpty.textContent = 'Dia fora do expediente configurado (segunda a sexta).';
    return;
  }

  if (!state.pcpEfficiencySnapshot) {
    elements.pcpEffEmpty.style.display = 'block';
    elements.pcpEffEmpty.textContent = 'Clique em Atualizar dados para gerar o painel.';
    renderPcpEffMachineFilter();
    return;
  }

  elements.pcpEffEmpty.style.display = 'none';
  const allMachineCards = state.pcpEfficiencySnapshot.machines || [];
  const selectedIds = new Set(getPcpSelectedMachineIds(allMachineCards));
  const machineCards = allMachineCards.filter((item) => selectedIds.has(item.listType));
  state.pcpEfficiencySelectedMachines = Array.from(selectedIds);
  renderPcpEffMachineFilter();

  if (!machineCards.length) {
    elements.pcpEffEmpty.style.display = 'block';
    elements.pcpEffEmpty.textContent = 'Selecione ao menos uma maquina para visualizar o painel.';
    return;
  }

  const overall = {
    totalOrdens: machineCards.reduce((acc, item) => acc + (item.totalOrdens || 0), 0),
    ordensFeitas: machineCards.reduce((acc, item) => acc + (item.ordensFeitas || 0), 0),
    totalPecasPlanejadas: machineCards.reduce((acc, item) => acc + (item.totalPecasPlanejadas || 0), 0),
    totalPecasProduzidas: machineCards.reduce((acc, item) => acc + (item.totalPecasProduzidas || 0), 0)
  };
  const overallOrdensPercent = overall.totalOrdens > 0 ? (overall.ordensFeitas / overall.totalOrdens) * 100 : 0;
  const overallPecasPercent = overall.totalPecasPlanejadas > 0
    ? (overall.totalPecasProduzidas / overall.totalPecasPlanejadas) * 100
    : 0;
  const overallActualBase = overallOrdensPercent;
  const overallStatus = evaluatePcpEffStatus(overallActualBase, planning.plannedPercent);
  const overallCard = document.createElement('article');
  overallCard.className = 'pcp-eff-card pcp-eff-overall-card';
  overallCard.classList.add(`pcp-eff-card-${overallStatus}`);

  const overallHeader = document.createElement('div');
  overallHeader.className = 'pcp-eff-card-header';
  const overallTitle = document.createElement('h3');
  overallTitle.textContent = 'Visão Geral';
  overallHeader.appendChild(overallTitle);
  overallHeader.appendChild(
    createPcpEffStatusVisual(
      overallStatus,
      'Eficiência por ordens',
      overallActualBase,
      planning.plannedPercent
    )
  );
  overallCard.appendChild(overallHeader);

  const overallMetrics = document.createElement('div');
  overallMetrics.className = 'pcp-eff-metrics-grid';
  overallMetrics.appendChild(
    createPcpEffMetricBlock({
      title: 'Ordens',
      total: overall.totalOrdens || 0,
      done: overall.ordensFeitas || 0,
      expected: (overall.totalOrdens || 0) * planning.progress,
      actualPercent: overallOrdensPercent,
      plannedPercent: planning.plannedPercent
    })
  );
  overallMetrics.appendChild(
    createPcpEffMetricBlock({
      title: 'Peças',
      total: overall.totalPecasPlanejadas || 0,
      done: overall.totalPecasProduzidas || 0,
      expected: (overall.totalPecasPlanejadas || 0) * planning.progress,
      actualPercent: overallPecasPercent,
      plannedPercent: planning.plannedPercent
    })
  );
  overallCard.appendChild(overallMetrics);

  const overallCharts = document.createElement('div');
  overallCharts.className = 'pcp-eff-charts-grid';
  overallCharts.appendChild(
    createPcpEffColumnChart(
      'Ordens feitas x planejadas no horário',
      overall.ordensFeitas || 0,
      (overall.totalOrdens || 0) * planning.progress,
      'Feitas',
      'Planejadas'
    )
  );
  overallCharts.appendChild(
    createPcpEffColumnChart(
      'Peças feitas x planejadas no horário',
      overall.totalPecasProduzidas || 0,
      (overall.totalPecasPlanejadas || 0) * planning.progress,
      'Feitas',
      'Planejadas'
    )
  );
  overallCard.appendChild(overallCharts);
  elements.pcpEffOverall.appendChild(overallCard);

  machineCards.forEach((machine) => {
    const totalOrdens = machine.totalOrdens || 0;
    const ordensFeitas = machine.ordensFeitas || 0;
    const totalPecas = machine.totalPecasPlanejadas || 0;
    const pecasFeitas = machine.totalPecasProduzidas || 0;
    const ordensPercent = totalOrdens > 0 ? (ordensFeitas / totalOrdens) * 100 : 0;
    const pecasPercent = totalPecas > 0 ? (pecasFeitas / totalPecas) * 100 : 0;
    const basis = getPcpMachineBasis(machine);
    const basisPercent = basis === 'ordens' ? ordensPercent : pecasPercent;
    const status = evaluatePcpEffStatus(basisPercent, planning.plannedPercent);

    const card = document.createElement('article');
    card.className = 'pcp-eff-card';
    card.classList.add(`pcp-eff-card-${status}`);

    const header = document.createElement('div');
    header.className = 'pcp-eff-card-header';
    const title = document.createElement('h3');
    title.textContent = machine.machineName || machine.listType.toUpperCase();
    header.appendChild(title);
    header.appendChild(
      createPcpEffStatusVisual(
        status,
        basis === 'ordens' ? 'Eficiência por ordens' : 'Eficiência por peças',
        basisPercent,
        planning.plannedPercent
      )
    );
    card.appendChild(header);

    const metrics = document.createElement('div');
    metrics.className = 'pcp-eff-metrics-grid';
    metrics.appendChild(
      createPcpEffMetricBlock({
        title: 'Ordens',
        total: totalOrdens,
        done: ordensFeitas,
        expected: totalOrdens * planning.progress,
        actualPercent: ordensPercent,
        plannedPercent: planning.plannedPercent
      })
    );
    metrics.appendChild(
      createPcpEffMetricBlock({
        title: 'Peças',
        total: totalPecas,
        done: pecasFeitas,
        expected: totalPecas * planning.progress,
        actualPercent: pecasPercent,
        plannedPercent: planning.plannedPercent
      })
    );
    card.appendChild(metrics);

    const charts = document.createElement('div');
    charts.className = 'pcp-eff-charts-grid';
    charts.appendChild(
      createPcpEffColumnChart(
        'Ordens feitas x planejadas no horário',
        ordensFeitas,
        totalOrdens * planning.progress,
        'Feitas',
        'Planejadas'
      )
    );
    charts.appendChild(
      createPcpEffColumnChart(
        'Peças feitas x planejadas no horário',
        pecasFeitas,
        totalPecas * planning.progress,
        'Feitas',
        'Planejadas'
      )
    );
    card.appendChild(charts);

    if (machine.missing) {
      const missingNote = document.createElement('p');
      missingNote.className = 'pcp-eff-missing';
      missingNote.textContent = 'Lista da máquina não encontrada para esta data.';
      card.appendChild(missingNote);
    }

    elements.pcpEffGrid.appendChild(card);
  });
}

function updatePcpEffTimeButtons() {
  const chips = [elements.pcpEffTimeAuto, elements.pcpEffTime10, elements.pcpEffTime12, elements.pcpEffTime15]
    .filter(Boolean);
  chips.forEach((chip) => {
    const mode = chip.dataset.hourMode || 'auto';
    chip.classList.toggle('active', mode === state.pcpEfficiencyHourMode);
  });
}

async function loadPcpEfficiencySnapshot() {
  const date = ensurePcpEffDateValue();
  state.pcpEfficiencySnapshot = await window.api.loadPcpEfficiency({ date });
  state.pcpEfficiencySelectedMachines = (state.pcpEfficiencySnapshot.machines || []).map((item) => item.listType);
}

async function openPcpEfficiency() {
  if (!state.pcpEfficiencySnapshot) {
    await loadPcpEfficiencySnapshot();
  }
  setPcpView('eff');
  updatePcpEffTimeButtons();
  renderPcpEfficiency();
}

function ensurePcpDashboardDateValue() {
  if (elements.pcpDashboardDate && elements.pcpDashboardDate.value) {
    return elements.pcpDashboardDate.value;
  }
  const baseDate = elements.dateInput.value || getTodayYmdLocal();
  if (elements.pcpDashboardDate) {
    elements.pcpDashboardDate.value = baseDate;
  }
  return baseDate;
}

function ensurePcpMoldagemDateValue() {
  return ensurePcpDashboardDateValue();
}

function formatWeekdayAndDatePtBr(dateValue) {
  const text = String(dateValue || '').trim();
  const parsed = new Date(`${text}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return formatYmdToBr(text);
  }
  const weekdayRaw = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(parsed);
  const weekday = weekdayRaw.charAt(0).toUpperCase() + weekdayRaw.slice(1);
  return `${weekday} - ${formatYmdToBr(text)}`;
}

function formatCurrencyBr(value) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 2 }).format(Number(value || 0));
}

function formatKg(value) {
  return `${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 3 }).format(Number(value || 0))} Kg`;
}

function formatDurationFromSeconds(totalSeconds) {
  const seconds = Math.max(0, Math.floor(Number(totalSeconds || 0)));
  if (!seconds) {
    return '0m';
  }
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  if (hours <= 0) {
    return `${minutes}m`;
  }
  return `${hours}h ${String(minutes).padStart(2, '0')}m`;
}

function formatDecimalMax3(value) {
  return new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 3 }).format(Number(value || 0));
}

function createDashboardPdfSectionHeader(title, dateValue, breakBefore = false) {
  const header = document.createElement('div');
  header.className = `pcp-dashboard-pdf-header pcp-section-pdf-header${breakBefore ? ' pcp-section-pdf-header-break' : ''}`;

  const logo = document.createElement('img');
  logo.className = 'pcp-dashboard-pdf-logo';
  const basePdfLogo = document.getElementById('pcp-dashboard-pdf-logo');
  logo.src = (basePdfLogo && basePdfLogo.getAttribute('src')) || '../../img/LogoParts.png';
  logo.alt = 'PartsSeals';

  const h2 = document.createElement('h2');
  h2.textContent = `${title} - ${formatWeekdayAndDatePtBr(dateValue)}`;

  header.appendChild(logo);
  header.appendChild(h2);
  return header;
}

function createDashboardKpiCard(label, value, secondary = '') {
  const card = document.createElement('article');
  card.className = 'pcp-dashboard-kpi';
  const p = document.createElement('p');
  p.textContent = label;
  const strong = document.createElement('strong');
  strong.textContent = value;
  card.appendChild(p);
  card.appendChild(strong);
  if (secondary) {
    const sub = document.createElement('span');
    sub.textContent = secondary;
    card.appendChild(sub);
  }
  return card;
}

function createDashboardBarChart({ title, data, colorResolver, valueFormatter, cardClass = '', labelMaxLen = 20 }) {
  const wrap = document.createElement('article');
  wrap.className = `pcp-dashboard-chart-card${cardClass ? ` ${cardClass}` : ''}`;
  if (data.length >= 14) {
    wrap.classList.add('pcp-dashboard-chart-card-dense');
  } else if (data.length >= 10) {
    wrap.classList.add('pcp-dashboard-chart-card-compact');
  }
  const h4 = document.createElement('h4');
  h4.textContent = title;
  wrap.appendChild(h4);

  const maxValue = Math.max(1, ...data.map((item) => Number(item.value || 0)));
  const list = document.createElement('div');
  list.className = 'pcp-dashboard-vcols';
  list.style.setProperty('--vcols-count', String(Math.max(1, data.length)));

  data.forEach((item) => {
    const col = document.createElement('div');
    col.className = 'pcp-dashboard-vcol-item';

    const value = document.createElement('span');
    value.className = 'pcp-dashboard-vcol-value';
    value.textContent = valueFormatter ? valueFormatter(item.value) : formatInteger(item.value);

    const barWrap = document.createElement('div');
    barWrap.className = 'pcp-dashboard-vcol-bar-wrap';
    const bar = document.createElement('div');
    bar.className = 'pcp-dashboard-vcol-bar';
    bar.style.height = `${Math.max(2, (Number(item.value || 0) / maxValue) * 100)}%`;
    bar.style.background = colorResolver ? colorResolver(item) : 'var(--primary)';
    barWrap.appendChild(bar);

    const label = document.createElement('span');
    label.className = 'pcp-dashboard-vcol-label';
    label.textContent = abbreviateChartLabel(String(item.name || ''), labelMaxLen);
    label.title = String(item.name || '');

    col.appendChild(value);
    col.appendChild(barWrap);
    col.appendChild(label);
    list.appendChild(col);
  });

  wrap.appendChild(list);
  return wrap;
}

function createDashboardDualChart(title, madeLabel, madeValue, plannedLabel, plannedValue) {
  return createDashboardBarChart({
    title,
    data: [
      { name: madeLabel, value: madeValue },
      { name: plannedLabel, value: plannedValue }
    ],
    colorResolver: (item) => normalizeForFilter(item.name).includes('FEIT') ? '#0ea5e9' : '#94a3b8',
    valueFormatter: (value) => formatInteger(value)
  });
}

function createDashboardDailySplitChart(charts) {
  const pieces = charts && charts.piecesDaily ? charts.piecesDaily : { feitas: 0, planejadas: 0 };
  const ops = charts && charts.opsDaily ? charts.opsDaily : { feitas: 0, planejadas: 0 };
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-dashboard-chart-card-daily-split';
  const h4 = document.createElement('h4');
  h4.textContent = 'Diárias (Peças e OPs)';
  wrap.appendChild(h4);

  const legend = document.createElement('div');
  legend.className = 'pcp-dashboard-daily-legend';
  legend.innerHTML = '<span><i style="background:#16a34a"></i>Feitas</span><span><i style="background:#dc2626"></i>Planejadas</span>';
  wrap.appendChild(legend);

  const grid = document.createElement('div');
  grid.className = 'pcp-dashboard-daily-split-grid';

  grid.appendChild(createDashboardBarChart({
    title: 'Peças Diárias',
    data: [
      { name: 'Feitas', value: pieces.feitas },
      { name: 'Planejadas', value: pieces.planejadas }
    ],
    colorResolver: (item) => normalizeForFilter(item.name).includes('FEIT') ? '#16a34a' : '#dc2626',
    valueFormatter: (v) => formatInteger(v),
    cardClass: 'pcp-dashboard-chart-card-mini',
    labelMaxLen: 14
  }));

  grid.appendChild(createDashboardBarChart({
    title: 'OPs Diárias',
    data: [
      { name: 'Feitas', value: ops.feitas },
      { name: 'Planejadas', value: ops.planejadas }
    ],
    colorResolver: (item) => normalizeForFilter(item.name).includes('FEIT') ? '#16a34a' : '#dc2626',
    valueFormatter: (v) => formatInteger(v),
    cardClass: 'pcp-dashboard-chart-card-mini',
    labelMaxLen: 14
  }));

  wrap.appendChild(grid);
  return wrap;
}

function abbreviateChartLabel(value, maxLen = 20) {
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) {
    return 'SEM NOME';
  }

  const dictionary = [
    [/\bASSISTENCIA\b/g, 'ASSIST.'],
    [/\bTECNICA\b/g, 'TEC.'],
    [/\bINDUSTRIA\b/g, 'IND.'],
    [/\bINDUSTRIAL\b/g, 'IND.'],
    [/\bCOMERCIO\b/g, 'COM.'],
    [/\bSERVICOS\b/g, 'SERV.'],
    [/\bSOLUCOES\b/g, 'SOL.'],
    [/\bSISTEMAS\b/g, 'SIST.'],
    [/\bMECANICA\b/g, 'MEC.'],
    [/\bENGENHARIA\b/g, 'ENG.'],
    [/\bCOMPONENTES\b/g, 'COMP.'],
    [/\bACABAMENTO\b/g, 'ACAB.'],
    [/\bMAQUINAS\b/g, 'MQS.'],
    [/\bNORTESTE\b/g, 'NE'],
    [/\bLTDA\b/g, 'LTDA'],
    [/\bEIRELI\b/g, 'EIR.'],
    [/\bS\.A\b/g, 'SA'],
    [/\bS\/A\b/g, 'SA'],
    [/\bDO\b/g, ''],
    [/\bDA\b/g, ''],
    [/\bDE\b/g, ''],
    [/\bDOS\b/g, ''],
    [/\bDAS\b/g, '']
  ];

  let text = raw;
  dictionary.forEach(([pattern, replace]) => {
    text = text.replace(pattern, replace);
  });
  text = text.replace(/\s+/g, ' ').trim();

  if (text.length <= maxLen) {
    return text;
  }

  const parts = text.split(' ').filter(Boolean);
  if (parts.length === 1) {
    return text.length > maxLen ? `${text.slice(0, Math.max(3, maxLen - 1))}…` : text;
  }

  let compact = parts
    .map((part) => (part.length > 10 ? `${part.slice(0, 8)}.` : part))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (compact.length <= maxLen) {
    return compact;
  }

  compact = parts
    .map((part) => {
      if (part.length <= 4) {
        return part;
      }
      return `${part.slice(0, 3)}.`;
    })
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (compact.length <= maxLen) {
    return compact;
  }
  return `${compact.slice(0, Math.max(3, maxLen - 1))}…`;
}

function formatDashboardCompanyLabel(name) {
  const rawOriginal = String(name || '').trim();
  if (!rawOriginal) {
    return 'SEM NOME';
  }
  const raw = abbreviateChartLabel(rawOriginal, 24);
  if (raw.length <= 16) {
    return raw;
  }

  const parts = raw.split(/\s+/).filter(Boolean);
  if (parts.length <= 1) {
    return raw.length > 22 ? `${raw.slice(0, 21)}…` : raw;
  }

  let line1 = '';
  let line2 = '';
  parts.forEach((part) => {
    const nextLine1 = `${line1} ${part}`.trim();
    if (!line2 && nextLine1.length <= 16) {
      line1 = nextLine1;
      return;
    }
    line2 = `${line2} ${part}`.trim();
  });

  if (!line1) {
    line1 = parts[0];
    line2 = parts.slice(1).join(' ');
  }
  if (line2.length > 16) {
    line2 = `${line2.slice(0, 15)}…`;
  }
  return line2 ? `${line1}\n${line2}` : line1;
}

function createDashboardHorizontalCompanyChart({ title, data, colorResolver, valueFormatter }) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-dashboard-chart-card-wide';
  if (data.length >= 18) {
    wrap.classList.add('pcp-dashboard-chart-card-ultra');
  } else if (data.length >= 13) {
    wrap.classList.add('pcp-dashboard-chart-card-dense');
  }
  const h4 = document.createElement('h4');
  h4.textContent = title;
  wrap.appendChild(h4);

  const maxValue = Math.max(1, ...data.map((item) => Number(item.value || 0)));
  const list = document.createElement('div');
  list.className = 'pcp-dashboard-hcols';
  list.style.setProperty('--hcols-count', String(Math.max(1, data.length)));

  data.forEach((item) => {
    const col = document.createElement('div');
    col.className = 'pcp-dashboard-hcol-item';

    const value = document.createElement('span');
    value.className = 'pcp-dashboard-hcol-value';
    value.textContent = valueFormatter ? valueFormatter(item.value) : formatInteger(item.value);

    const barWrap = document.createElement('div');
    barWrap.className = 'pcp-dashboard-hcol-bar-wrap';
    const bar = document.createElement('div');
    bar.className = 'pcp-dashboard-hcol-bar';
    bar.style.height = `${Math.max(2, (Number(item.value || 0) / maxValue) * 100)}%`;
    bar.style.background = colorResolver ? colorResolver(item) : 'var(--primary)';
    barWrap.appendChild(bar);

    const label = document.createElement('span');
    label.className = 'pcp-dashboard-hcol-label';
    label.textContent = formatDashboardCompanyLabel(item.name);
    label.title = String(item.name || '');

    col.appendChild(value);
    col.appendChild(barWrap);
    col.appendChild(label);
    list.appendChild(col);
  });

  wrap.appendChild(list);
  return wrap;
}

function createDashboardPlaceholderChart(title, message) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card';
  const h4 = document.createElement('h4');
  h4.textContent = title;
  wrap.appendChild(h4);
  const p = document.createElement('p');
  p.className = 'pcp-cnc-placeholder';
  p.textContent = message;
  wrap.appendChild(p);
  return wrap;
}

function createCncComparisonChart({ title, data, doneKey, plannedKey, valueFormatter, colorResolver }) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-cnc-compare-card';
  const h4 = document.createElement('h4');
  h4.textContent = title;
  wrap.appendChild(h4);

  const maxValue = Math.max(1, ...data.map((item) => Math.max(Number(item[doneKey] || 0), Number(item[plannedKey] || 0))));
  const list = document.createElement('div');
  list.className = 'pcp-cnc-compare-list';

  data.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'pcp-cnc-compare-row';

    const label = document.createElement('span');
    label.className = 'pcp-cnc-compare-label';
    label.textContent = item.machine;

    const bars = document.createElement('div');
    bars.className = 'pcp-cnc-compare-bars';

    const done = document.createElement('div');
    done.className = 'pcp-cnc-compare-bar done';
    done.style.width = `${Math.max(2, (Number(item[doneKey] || 0) / maxValue) * 100)}%`;
    done.style.background = colorResolver ? colorResolver(item) : '#16a34a';
    done.title = `Produzidas: ${valueFormatter(item[doneKey] || 0)}`;

    const planned = document.createElement('div');
    planned.className = 'pcp-cnc-compare-bar planned';
    planned.style.width = `${Math.max(2, (Number(item[plannedKey] || 0) / maxValue) * 100)}%`;
    planned.style.background = colorResolver ? colorResolver(item) : '#dc2626';
    planned.style.opacity = '0.45';
    planned.title = `Planejadas: ${valueFormatter(item[plannedKey] || 0)}`;

    bars.appendChild(done);
    bars.appendChild(planned);

    const value = document.createElement('span');
    value.className = 'pcp-cnc-compare-value';
    value.textContent = `${valueFormatter(item[doneKey] || 0)} / ${valueFormatter(item[plannedKey] || 0)}`;

    row.appendChild(label);
    row.appendChild(bars);
    row.appendChild(value);
    list.appendChild(row);
  });

  wrap.appendChild(list);
  return wrap;
}

function createCncVerticalComparisonChart({ title, data, doneKey, plannedKey, valueFormatter, colorResolver }) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-cnc-vertical-card';
  const h4 = document.createElement('h4');
  h4.textContent = title;
  wrap.appendChild(h4);
  const legend = document.createElement('div');
  legend.className = 'pcp-cnc-vertical-legend';
  legend.textContent = 'Escuro: Produzidas | Claro: Planejadas';
  wrap.appendChild(legend);

  const maxValue = Math.max(
    1,
    ...data.map((item) => Math.max(Number(item[doneKey] || 0), Number(item[plannedKey] || 0)))
  );
  const grid = document.createElement('div');
  grid.className = 'pcp-cnc-vertical-grid';

  data.forEach((item) => {
    const col = document.createElement('div');
    col.className = 'pcp-cnc-vertical-col';

    const machineLabel = document.createElement('span');
    machineLabel.className = 'pcp-cnc-vertical-machine';
    machineLabel.textContent = item.machine;

    const barsWrap = document.createElement('div');
    barsWrap.className = 'pcp-cnc-vertical-bars';

    const doneWrap = document.createElement('div');
    doneWrap.className = 'pcp-cnc-vertical-bar-wrap';
    const doneValue = document.createElement('span');
    doneValue.className = 'pcp-cnc-vertical-bar-value done';
    doneValue.textContent = valueFormatter(item[doneKey] || 0);
    const done = document.createElement('div');
    done.className = 'pcp-cnc-vertical-bar done';
    done.style.height = `${Math.max(2, (Number(item[doneKey] || 0) / maxValue) * 100)}%`;
    done.style.background = colorResolver ? colorResolver(item) : '#16a34a';
    done.title = `Produzidas: ${valueFormatter(item[doneKey] || 0)}`;
    doneWrap.appendChild(doneValue);
    doneWrap.appendChild(done);

    const plannedWrap = document.createElement('div');
    plannedWrap.className = 'pcp-cnc-vertical-bar-wrap';
    const plannedValue = document.createElement('span');
    plannedValue.className = 'pcp-cnc-vertical-bar-value planned';
    plannedValue.textContent = valueFormatter(item[plannedKey] || 0);
    const planned = document.createElement('div');
    planned.className = 'pcp-cnc-vertical-bar planned';
    planned.style.height = `${Math.max(2, (Number(item[plannedKey] || 0) / maxValue) * 100)}%`;
    planned.style.background = colorResolver ? colorResolver(item) : '#dc2626';
    planned.style.opacity = '0.45';
    planned.title = `Planejadas: ${valueFormatter(item[plannedKey] || 0)}`;
    plannedWrap.appendChild(plannedValue);
    plannedWrap.appendChild(planned);

    barsWrap.appendChild(doneWrap);
    barsWrap.appendChild(plannedWrap);

    const values = document.createElement('span');
    values.className = 'pcp-cnc-vertical-values';
    values.textContent = `Prod ${valueFormatter(item[doneKey] || 0)} | Plan ${valueFormatter(item[plannedKey] || 0)}`;

    col.appendChild(machineLabel);
    col.appendChild(barsWrap);
    col.appendChild(values);
    grid.appendChild(col);
  });

  wrap.appendChild(grid);
  return wrap;
}

function createCncContributionChart(data) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-cnc-pie-card';
  const h4 = document.createElement('h4');
  h4.textContent = 'Percentual de contribuição';
  wrap.appendChild(h4);

  const safe = (data || []).map((item, idx) => ({
    ...item,
    color: CNC_MACHINE_COLOR_MAP[item.machine] || ['#2563eb', '#16a34a', '#f59e0b', '#8b5cf6', '#ef4444'][idx % 5],
    pct: Number(item.contributionPercent || 0)
  }));
  let start = 0;
  const slices = [];
  const stops = safe.map((item) => {
    const sliceStart = start;
    const end = start + item.pct;
    const piece = `${item.color} ${start}% ${end}%`;
    slices.push({ ...item, start: sliceStart, end, mid: sliceStart + (item.pct / 2) });
    start = end;
    return piece;
  });
  const pie = document.createElement('div');
  pie.className = 'pcp-cnc-pie';
  pie.style.background = stops.length
    ? `conic-gradient(${stops.join(', ')})`
    : 'conic-gradient(#cbd5e1 0% 100%)';

  slices
    .filter((slice) => Number(slice.pct || 0) >= 6)
    .forEach((slice) => {
      const angle = ((slice.mid / 100) * (Math.PI * 2)) - (Math.PI / 2);
      const radius = 34;
      const x = 50 + (Math.cos(angle) * radius);
      const y = 50 + (Math.sin(angle) * radius);
      const label = document.createElement('span');
      label.className = 'pcp-cnc-pie-slice-label';
      label.textContent = formatPercent(slice.pct);
      label.style.left = `${x}%`;
      label.style.top = `${y}%`;
      pie.appendChild(label);
    });

  slices
    .filter((slice) => Number(slice.pct || 0) > 0 && Number(slice.pct || 0) < 6)
    .forEach((slice) => {
      const angle = ((slice.mid / 100) * (Math.PI * 2)) - (Math.PI / 2);
      const radius = 58;
      const x = 50 + (Math.cos(angle) * radius);
      const y = 50 + (Math.sin(angle) * radius);
      const callout = document.createElement('span');
      callout.className = 'pcp-cnc-pie-slice-callout';
      callout.textContent = formatPercent(slice.pct);
      callout.style.left = `${x}%`;
      callout.style.top = `${y}%`;
      pie.appendChild(callout);
    });

  wrap.appendChild(pie);

  const legend = document.createElement('div');
  legend.className = 'pcp-cnc-pie-legend';
  safe.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'pcp-cnc-pie-legend-row';
    const dot = document.createElement('i');
    dot.style.background = item.color;
    const text = document.createElement('span');
    text.textContent = `${item.machine}: ${formatPercent(item.pct)}`;
    row.appendChild(dot);
    row.appendChild(text);
    legend.appendChild(row);
  });
  wrap.appendChild(legend);
  return wrap;
}

function createRncSectorPieChart(data) {
  const wrap = document.createElement('article');
  wrap.className = 'pcp-dashboard-chart-card pcp-cnc-pie-card';
  const h4 = document.createElement('h4');
  h4.textContent = 'RNC por Setor (%)';
  wrap.appendChild(h4);

  const safe = (data || []).map((item, idx) => {
    const sectorName = String(item.sector || item.name || '').trim().toUpperCase() || `SETOR ${idx + 1}`;
    return {
      sector: sectorName,
      count: Number(item.count || 0),
      pct: Number(item.pct || 0),
      color: RNC_SECTOR_COLOR_MAP[normalizeForFilter(sectorName)] || ['#2563eb', '#16a34a', '#f59e0b', '#8b5cf6', '#ef4444'][idx % 5]
    };
  });

  let start = 0;
  const slices = [];
  const stops = safe.map((item) => {
    const sliceStart = start;
    const end = start + item.pct;
    const piece = `${item.color} ${start}% ${end}%`;
    slices.push({ ...item, start: sliceStart, end, mid: sliceStart + (item.pct / 2) });
    start = end;
    return piece;
  });

  const pie = document.createElement('div');
  pie.className = 'pcp-cnc-pie';
  pie.style.background = stops.length
    ? `conic-gradient(${stops.join(', ')})`
    : 'conic-gradient(#cbd5e1 0% 100%)';

  slices
    .filter((slice) => Number(slice.pct || 0) >= 6)
    .forEach((slice) => {
      const angle = ((slice.mid / 100) * (Math.PI * 2)) - (Math.PI / 2);
      const radius = 34;
      const x = 50 + (Math.cos(angle) * radius);
      const y = 50 + (Math.sin(angle) * radius);
      const label = document.createElement('span');
      label.className = 'pcp-cnc-pie-slice-label';
      label.textContent = formatPercent(slice.pct);
      label.style.left = `${x}%`;
      label.style.top = `${y}%`;
      pie.appendChild(label);
    });

  slices
    .filter((slice) => Number(slice.pct || 0) > 0 && Number(slice.pct || 0) < 6)
    .forEach((slice) => {
      const angle = ((slice.mid / 100) * (Math.PI * 2)) - (Math.PI / 2);
      const radius = 58;
      const x = 50 + (Math.cos(angle) * radius);
      const y = 50 + (Math.sin(angle) * radius);
      const callout = document.createElement('span');
      callout.className = 'pcp-cnc-pie-slice-callout';
      callout.textContent = formatPercent(slice.pct);
      callout.style.left = `${x}%`;
      callout.style.top = `${y}%`;
      pie.appendChild(callout);
    });

  wrap.appendChild(pie);

  const legend = document.createElement('div');
  legend.className = 'pcp-cnc-pie-legend';
  safe.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'pcp-cnc-pie-legend-row';
    const dot = document.createElement('i');
    dot.style.background = item.color;
    const text = document.createElement('span');
    text.textContent = `${item.sector}: ${formatPercent(item.pct)} (${formatInteger(item.count)})`;
    row.appendChild(dot);
    row.appendChild(text);
    legend.appendChild(row);
  });
  wrap.appendChild(legend);
  return wrap;
}

function renderPcpDashboard() {
  if (!elements.pcpDashboardOverview || !elements.pcpDashboardCharts || !elements.pcpDashboardEmpty || !elements.pcpDashboardTitle) {
    return;
  }
  if (state.pcpView !== 'dashboard') {
    return;
  }

  elements.pcpDashboardOverview.innerHTML = '';
  elements.pcpDashboardCharts.innerHTML = '';
  if (elements.pcpMoldagemOverview) {
    elements.pcpMoldagemOverview.innerHTML = '';
  }
  if (elements.pcpMoldagemCharts) {
    elements.pcpMoldagemCharts.innerHTML = '';
  }
  if (elements.pcpMoldagemEmpty) {
    elements.pcpMoldagemEmpty.style.display = 'none';
  }
  if (elements.pcpCncOverview) {
    elements.pcpCncOverview.innerHTML = '';
  }
  if (elements.pcpCncCharts) {
    elements.pcpCncCharts.innerHTML = '';
  }
  if (elements.pcpCncEmpty) {
    elements.pcpCncEmpty.style.display = 'none';
  }
  if (elements.pcpRncOverview) {
    elements.pcpRncOverview.innerHTML = '';
  }
  if (elements.pcpRncCharts) {
    elements.pcpRncCharts.innerHTML = '';
  }
  if (elements.pcpRncEmpty) {
    elements.pcpRncEmpty.style.display = 'none';
  }

  const dateValue = ensurePcpDashboardDateValue();
  if (!state.pcpDashboardSnapshot) {
    elements.pcpDashboardEmpty.style.display = 'block';
    elements.pcpDashboardEmpty.textContent = 'Clique em Atualizar dados para gerar o dashboard.';
    elements.pcpDashboardTitle.textContent = `Visão geral de produção - ${dateValue}`;
    if (elements.pcpMoldagemEmpty) {
      elements.pcpMoldagemEmpty.style.display = 'block';
      elements.pcpMoldagemEmpty.textContent = 'Clique em Atualizar dados para gerar o painel de moldagem.';
    }
    if (elements.pcpCncEmpty) {
      elements.pcpCncEmpty.style.display = 'block';
      elements.pcpCncEmpty.textContent = 'Clique em Atualizar dados para gerar o painel de usinagem CNC.';
    }
    if (elements.pcpRncEmpty) {
      elements.pcpRncEmpty.style.display = 'block';
      elements.pcpRncEmpty.textContent = 'Sem dados de RNC para a data selecionada.';
    }
    return;
  }

  const snapshot = state.pcpDashboardSnapshot;
  elements.pcpDashboardEmpty.style.display = 'none';
  elements.pcpDashboardTitle.textContent = `${snapshot.title} | Fonte: ${snapshot.sourceSheet}`;

  const k = snapshot.kpis || {};
  const kpiRow1 = document.createElement('div');
  kpiRow1.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRow1.appendChild(createDashboardKpiCard('Peças Feitas', formatInteger(k.pecasFeitas)));
  kpiRow1.appendChild(createDashboardKpiCard('Peças Planejadas', formatInteger(k.pecasPlanejadas)));
  kpiRow1.appendChild(createDashboardKpiCard('N OP Feitas', formatInteger(k.opFeitas)));
  kpiRow1.appendChild(createDashboardKpiCard('N OP Planejadas', formatInteger(k.opPlanejadas)));

  const kpiRow2 = document.createElement('div');
  kpiRow2.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRow2.appendChild(createDashboardKpiCard('Efic Total (OPs)', formatPercent(k.eficOps)));
  kpiRow2.appendChild(createDashboardKpiCard('OP Geradas', formatInteger(k.opGeradas)));
  kpiRow2.appendChild(createDashboardKpiCard('% Teflon (OPs)', formatPercent(k.teflonOpsPercent)));
  kpiRow2.appendChild(createDashboardKpiCard('% Teflon (Peças)', formatPercent(k.teflonPecasPercent)));

  const kpiRow3 = document.createElement('div');
  kpiRow3.className = 'pcp-dashboard-kpi-row cols-3';
  kpiRow3.appendChild(createDashboardKpiCard('Qntd RNC', formatInteger(k.rncQtd)));
  kpiRow3.appendChild(createDashboardKpiCard('Custo RNC', formatCurrencyBr(k.rncCost)));
  kpiRow3.appendChild(createDashboardKpiCard('Custo Total', formatCurrencyBr(k.totalCost)));

  elements.pcpDashboardOverview.appendChild(kpiRow1);
  elements.pcpDashboardOverview.appendChild(kpiRow2);
  elements.pcpDashboardOverview.appendChild(kpiRow3);

  if ((Number(k.rncQtd || 0) <= 0) && (Number(k.rncCost || 0) <= 0)) {
    const rncSnapshot = snapshot && snapshot.debug ? snapshot.debug.rnc : null;
    const rncDebug = rncSnapshot && rncSnapshot.debug ? rncSnapshot.debug : null;
    if (rncDebug) {
      const note = document.createElement('p');
      note.className = 'field-note';
      const err = rncDebug.error || '';
      const file = rncDebug.sourceFile || '';
      const sheet = rncDebug.sheetName || '';
      const expected = rncDebug.expected || '';
      const fiscalRoot = rncDebug.fiscalRoot || '';
      const rootOk = rncDebug.fiscalRootExists === true ? 'OK' : 'NOK';
      const expectedOk = rncDebug.expectedExists === true ? 'OK' : 'NOK';
      const headerIndex = Number.isFinite(rncDebug.headerIndex) ? ` | headerIndex: ${rncDebug.headerIndex + 1}` : '';
      const runtime = snapshot && snapshot.debug ? snapshot.debug.runtime : null;
      const runtimeInfo = runtime && (runtime.mainFile || runtime.appVersion)
        ? ` | main: ${runtime.mainFile || '-'} | ver: ${runtime.appVersion || '-'}`
        : '';

      note.textContent = err
        ? `RNC: sem dados (${err})`
        : `RNC: sem dados (arquivo: ${file || '-'} | aba: ${sheet || '-'}${headerIndex} | root: ${fiscalRoot || '-'} (${rootOk}) | esperado: ${expected || '-'} (${expectedOk})${runtimeInfo})`;
      if (err) {
        note.classList.add('error');
      }
      elements.pcpDashboardOverview.appendChild(note);
    }
  }

  const operatorColorMap = {
    GUSTAVO: '#16a34a',
    DARCI: '#0ea5e9',
    REBECA: '#f59e0b',
    PALOMA: '#a855f7'
  };
  const getOperatorColor = (item) => operatorColorMap[normalizeForFilter(item && item.name)] || '#22c55e';

  const charts = snapshot.charts || {};
  elements.pcpDashboardCharts.appendChild(createDashboardHorizontalCompanyChart({
    title: 'Qtd OP por Empresa',
    data: charts.opsByCompany || [],
    colorResolver: () => 'var(--primary)',
    valueFormatter: (v) => formatInteger(v)
  }));
  elements.pcpDashboardCharts.appendChild(createDashboardDailySplitChart(charts));
  elements.pcpDashboardCharts.appendChild(createDashboardBarChart({
    title: 'Acabamento - OPS por Operador',
    data: (charts.acabamentoByOperator || []).map((item) => ({ name: item.operador, value: item.ops || 0 })),
    colorResolver: (item) => getOperatorColor(item),
    valueFormatter: (v) => formatInteger(v)
  }));
  elements.pcpDashboardCharts.appendChild(createDashboardBarChart({
    title: 'Acabamento - Peças por Operador',
    data: (charts.acabamentoByOperator || []).map((item) => ({ name: item.operador, value: item.pecas || 0 })),
    colorResolver: (item) => getOperatorColor(item),
    valueFormatter: (v) => formatInteger(v)
  }));
  elements.pcpDashboardCharts.appendChild(createDashboardBarChart({
    title: 'Material Utilizado (Kg)',
    data: (charts.materials || []).map((item) => ({ name: item.name, value: item.kgUsed || 0 })),
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatKg(v),
    cardClass: 'pcp-dashboard-chart-card-material',
    labelMaxLen: 26
  }));
  elements.pcpDashboardCharts.appendChild(createDashboardBarChart({
    title: 'Material em Estoque',
    data: (charts.materials || []).map((item) => ({ name: item.name, value: item.estoque || 0 })),
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatDecimalMax3(v),
    cardClass: 'pcp-dashboard-chart-card-material pcp-dashboard-chart-card-stock',
    labelMaxLen: 26
  }));

  renderPcpMoldagem();
  renderPcpCnc();
  renderPcpRnc();
}

function renderPcpMoldagem() {
  if (!elements.pcpMoldagemOverview || !elements.pcpMoldagemCharts || !elements.pcpMoldagemEmpty || !elements.pcpMoldagemTitle) {
    return;
  }
  if (state.pcpView !== 'dashboard') {
    return;
  }

  elements.pcpMoldagemOverview.innerHTML = '';
  elements.pcpMoldagemCharts.innerHTML = '';

  const dateValue = ensurePcpMoldagemDateValue();
  elements.pcpMoldagemTitle.textContent = `MOLDAGEM - VISÃO GERAL - ${formatWeekdayAndDatePtBr(dateValue)}`;

  if (!state.pcpMoldagemSnapshot) {
    elements.pcpMoldagemEmpty.style.display = 'block';
    elements.pcpMoldagemEmpty.textContent = 'Clique em Atualizar dados para gerar o painel de moldagem.';
    return;
  }

  const snapshot = state.pcpMoldagemSnapshot;
  elements.pcpMoldagemEmpty.style.display = 'none';
  elements.pcpMoldagemOverview.appendChild(createDashboardPdfSectionHeader('MOLDAGEM - VISÃO GERAL', dateValue));

  const k = snapshot.kpis || {};
  const kpiRow1 = document.createElement('div');
  kpiRow1.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRow1.appendChild(createDashboardKpiCard('Efic Moldagem', formatPercent(k.eficMoldagem)));
  kpiRow1.appendChild(createDashboardKpiCard('KG Processados', formatKg(k.kgProcessados)));
  kpiRow1.appendChild(createDashboardKpiCard('Buchas Moldadas', formatDecimalMax3(k.buchasMoldadas)));
  kpiRow1.appendChild(createDashboardKpiCard('KGs em Refugo', formatKg(k.kgRefugo)));

  const kpiRow2 = document.createElement('div');
  kpiRow2.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRow2.appendChild(createDashboardKpiCard('Material de Maior Saída', k.materialMaiorSaida || '-'));
  kpiRow2.appendChild(createDashboardKpiCard('Material de Mais Buchas', k.materialMaisBuchas || '-'));
  kpiRow2.appendChild(createDashboardKpiCard('Custo Refugo', formatCurrencyBr(k.custoRefugo || 0)));
  kpiRow2.appendChild(createDashboardKpiCard('Data', formatYmdToBr(snapshot.date)));

  elements.pcpMoldagemOverview.appendChild(kpiRow1);
  elements.pcpMoldagemOverview.appendChild(kpiRow2);

  const charts = snapshot.charts || {};
  const refugoByMaterial = (charts.materials || [])
    .map((item) => ({ name: item.name, value: item.refugoKg || 0 }))
    .filter((item) => Number(item.value || 0) > 0);
  const opsByMaterial = (charts.materials || [])
    .map((item) => ({ name: item.name, value: item.ops || 0 }))
    .filter((item) => Number(item.value || 0) > 0);

  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Qntd Refugo por Material (Kg)',
    data: refugoByMaterial,
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatKg(v),
    labelMaxLen: 26
  }));
  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Qntd OP por Material',
    data: opsByMaterial,
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatInteger(v),
    labelMaxLen: 26
  }));
  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Buchas por Prensa',
    data: (charts.presses || []).map((item) => ({ name: item.name, value: item.buchas || 0 })),
    colorResolver: () => '#16a34a',
    valueFormatter: (v) => formatDecimalMax3(v)
  }));
  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'OP por Prensa',
    data: (charts.presses || []).map((item) => ({ name: item.name, value: item.ops || 0 })),
    colorResolver: () => '#0ea5e9',
    valueFormatter: (v) => formatInteger(v)
  }));
  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Material em Estoque',
    data: (charts.materials || []).map((item) => ({ name: item.name, value: item.estoque || 0 })),
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatDecimalMax3(v),
    cardClass: 'pcp-dashboard-chart-card-material',
    labelMaxLen: 26
  }));
  const moldagem2Header = createDashboardPdfSectionHeader('MOLDAGEM - PROCESSAMENTO', dateValue, true);
  moldagem2Header.classList.add('pcp-moldagem-page2-header');
  elements.pcpMoldagemCharts.appendChild(moldagem2Header);
  const totalOpsMaterial = (charts.materials || []).reduce((acc, item) => acc + Number(item.ops || 0), 0);
  const avgKgPerOp = totalOpsMaterial > 0 ? Number(k.kgProcessados || 0) / totalOpsMaterial : 0;
  const avgBuchasPerOp = totalOpsMaterial > 0 ? Number(k.buchasMoldadas || 0) / totalOpsMaterial : 0;
  const indiceRefugo = Number(k.kgProcessados || 0) > 0
    ? (Number(k.kgRefugo || 0) / Number(k.kgProcessados || 0)) * 100
    : 0;
  const materialMaiorRefugo = (charts.materials || [])
    .filter((item) => Number(item.refugoKg || 0) > 0)
    .reduce((best, item) => {
      if (!best || Number(item.refugoKg || 0) > Number(best.refugoKg || 0)) {
        return item;
      }
      return best;
    }, null);
  const materialLiderKg = (charts.materials || [])
    .reduce((best, item) => {
      if (!best || Number(item.kgUsed || 0) > Number(best.kgUsed || 0)) {
        return item;
      }
      return best;
    }, null);
  const participacaoMaterialLider = Number(k.kgProcessados || 0) > 0
    ? (Number((materialLiderKg && materialLiderKg.kgUsed) || 0) / Number(k.kgProcessados || 0)) * 100
    : 0;
  const buchasPorKg = Number(k.kgProcessados || 0) > 0
    ? Number(k.buchasMoldadas || 0) / Number(k.kgProcessados || 0)
    : 0;
  const topOperadorPeso = (charts.operators || [])
    .filter((item) => String(item.name || '').trim())
    .reduce((best, item) => {
      if (!best || Number(item.peso || 0) > Number(best.peso || 0)) {
        return item;
      }
      return best;
    }, null);

  const moldagemPage2KpiRow = document.createElement('div');
  moldagemPage2KpiRow.className = 'pcp-dashboard-kpi-row cols-4 pcp-moldagem-page2-kpis';
  moldagemPage2KpiRow.appendChild(createDashboardKpiCard('Ops Prensadas', formatInteger(totalOpsMaterial)));
  moldagemPage2KpiRow.appendChild(createDashboardKpiCard('Média KG por OP', formatKg(avgKgPerOp)));
  moldagemPage2KpiRow.appendChild(createDashboardKpiCard('Média Buchas por OP', formatDecimalMax3(avgBuchasPerOp)));
  moldagemPage2KpiRow.appendChild(createDashboardKpiCard('Índice de Refugo', formatPercent(indiceRefugo)));
  elements.pcpMoldagemCharts.appendChild(moldagemPage2KpiRow);

  const moldagemPage2KpiRowExtra = document.createElement('div');
  moldagemPage2KpiRowExtra.className = 'pcp-dashboard-kpi-row cols-4 pcp-moldagem-page2-kpis-extra';
  moldagemPage2KpiRowExtra.appendChild(
    createDashboardKpiCard(
      'Material com Maior Refugo',
      materialMaiorRefugo ? materialMaiorRefugo.name : '-',
      materialMaiorRefugo ? formatKg(materialMaiorRefugo.refugoKg || 0) : ''
    )
  );
  moldagemPage2KpiRowExtra.appendChild(
    createDashboardKpiCard(
      '% Participação Material Líder',
      formatPercent(participacaoMaterialLider),
      materialLiderKg ? `${materialLiderKg.name} (${formatKg(materialLiderKg.kgUsed || 0)})` : ''
    )
  );
  moldagemPage2KpiRowExtra.appendChild(
    createDashboardKpiCard('Buchas por Kg', formatDecimalMax3(buchasPorKg))
  );
  moldagemPage2KpiRowExtra.appendChild(
    createDashboardKpiCard(
      'Top Operador por Peso',
      topOperadorPeso ? topOperadorPeso.name : '-',
      topOperadorPeso ? formatKg(topOperadorPeso.peso || 0) : ''
    )
  );
  elements.pcpMoldagemCharts.appendChild(moldagemPage2KpiRowExtra);

  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Quantidade de Material Processado (Kg)',
    data: (charts.materials || []).map((item) => ({ name: item.name, value: item.kgUsed || 0 })),
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatKg(v),
    cardClass: 'pcp-dashboard-chart-card-material pcp-moldagem-processed-card',
    labelMaxLen: 26
  }));
  elements.pcpMoldagemCharts.appendChild(createDashboardBarChart({
    title: 'Quantidade de Buchas por Material',
    data: (charts.materials || [])
      .map((item) => ({ name: item.name, value: item.buchas || 0 }))
      .filter((item) => Number(item.value || 0) > 0),
    colorResolver: (item) => MATERIAL_COLOR_MAP[item.name] || '#64748b',
    valueFormatter: (v) => formatDecimalMax3(v),
    cardClass: 'pcp-dashboard-chart-card-material',
    labelMaxLen: 26
  }));
  const operatorsRow = document.createElement('div');
  operatorsRow.className = 'pcp-moldagem-operators-row';
  operatorsRow.appendChild(createDashboardBarChart({
    title: 'Eficiência Operador - OP',
    data: (charts.operators || []).map((item) => ({ name: item.name, value: item.ops || 0 })),
    colorResolver: () => '#4f46e5',
    valueFormatter: (v) => formatInteger(v)
  }));
  operatorsRow.appendChild(createDashboardBarChart({
    title: 'Eficiência Operador - Bucha',
    data: (charts.operators || []).map((item) => ({ name: item.name, value: item.buchas || 0 })),
    colorResolver: () => '#0891b2',
    valueFormatter: (v) => formatDecimalMax3(v)
  }));
  operatorsRow.appendChild(createDashboardBarChart({
    title: 'Eficiência Operador - Peso',
    data: (charts.operators || []).map((item) => ({ name: item.name, value: item.peso || 0 })),
    colorResolver: () => '#f97316',
    valueFormatter: (v) => formatKg(v)
  }));
  elements.pcpMoldagemCharts.appendChild(operatorsRow);
}

function renderPcpRnc() {
  if (!elements.pcpRncOverview || !elements.pcpRncCharts || !elements.pcpRncEmpty || !elements.pcpRncTitle) {
    return;
  }
  if (state.pcpView !== 'dashboard') {
    return;
  }

  elements.pcpRncOverview.innerHTML = '';
  elements.pcpRncCharts.innerHTML = '';

  const dateValue = ensurePcpDashboardDateValue();
  elements.pcpRncTitle.textContent = `RELATÓRIO RNC - ${formatWeekdayAndDatePtBr(dateValue)}`;

  const snapshot = state.pcpDashboardSnapshot;
  const report = snapshot && snapshot.charts ? snapshot.charts.rncReport : null;
  const k = report && report.kpis ? report.kpis : {};
  const charts = report && report.charts ? report.charts : {};
  const totalRnc = Number(k.rncQtd || 0);

  if (!totalRnc) {
    elements.pcpRncEmpty.style.display = 'block';
    elements.pcpRncEmpty.textContent = 'Sem dados de RNC para a data selecionada.';
    return;
  }

  elements.pcpRncEmpty.style.display = 'none';
  elements.pcpRncOverview.appendChild(createDashboardPdfSectionHeader('RELATÓRIO RNC', dateValue, true));

  const kpiRow1 = document.createElement('div');
  kpiRow1.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRow1.appendChild(createDashboardKpiCard('Qntd RNC', formatInteger(k.rncQtd || 0)));
  kpiRow1.appendChild(createDashboardKpiCard('Custo RNC', formatCurrencyBr(k.rncCost || 0)));
  kpiRow1.appendChild(createDashboardKpiCard('Custo médio RNC', formatCurrencyBr(k.avgCost || 0)));
  kpiRow1.appendChild(createDashboardKpiCard('Tempo perdido', formatDurationFromSeconds(k.segsPerdidos || 0)));

  const kpiRow2 = document.createElement('div');
  kpiRow2.className = 'pcp-dashboard-kpi-row cols-3';
  kpiRow2.appendChild(createDashboardKpiCard('Setor com mais RNC', String(k.topSectorByCount || '-')));
  kpiRow2.appendChild(createDashboardKpiCard('Setor com maior custo', String(k.topSectorByCost || '-')));
  kpiRow2.appendChild(createDashboardKpiCard('Maior motivo', String(k.topReason || '-')));

  elements.pcpRncOverview.appendChild(kpiRow1);
  elements.pcpRncOverview.appendChild(kpiRow2);

  elements.pcpRncCharts.appendChild(createRncSectorPieChart(charts.sectorPie || []));
  elements.pcpRncCharts.appendChild(createDashboardBarChart({
    title: 'Quantidade de RNC por Setor',
    data: charts.sectorCounts || [],
    colorResolver: (item) => RNC_SECTOR_COLOR_MAP[normalizeForFilter(item && item.name)] || '#64748b',
    valueFormatter: (v) => formatInteger(v)
  }));
  elements.pcpRncCharts.appendChild(createDashboardBarChart({
    title: 'Quantidade de RNC por Motivo',
    data: charts.reasonsCount || [],
    colorResolver: () => '#ef4444',
    valueFormatter: (v) => formatInteger(v),
    labelMaxLen: 36
  }));
  elements.pcpRncCharts.appendChild(createDashboardBarChart({
    title: 'Custo por Motivo',
    data: charts.reasonsCost || [],
    colorResolver: () => '#22c55e',
    valueFormatter: (v) => formatCurrencyBr(v),
    labelMaxLen: 36
  }));
  elements.pcpRncCharts.appendChild(createDashboardBarChart({
    title: 'Tempo perdido por Máquina',
    data: charts.machineTime || [],
    colorResolver: () => '#f59e0b',
    valueFormatter: (v) => formatDurationFromSeconds(v),
    labelMaxLen: 26
  }));
  elements.pcpRncCharts.appendChild(createDashboardBarChart({
    title: 'Quantidade de RNC por Material',
    data: charts.materialsCount || [],
    colorResolver: () => '#0ea5e9',
    valueFormatter: (v) => formatInteger(v),
    labelMaxLen: 26
  }));
}

function renderPcpCnc() {
  if (!elements.pcpCncOverview || !elements.pcpCncCharts || !elements.pcpCncEmpty || !elements.pcpCncTitle) {
    return;
  }
  if (state.pcpView !== 'dashboard') {
    return;
  }

  elements.pcpCncOverview.innerHTML = '';
  elements.pcpCncCharts.innerHTML = '';
  elements.pcpCncTitle.textContent = `USINAGEM MÁQUINAS CNC - ${formatWeekdayAndDatePtBr(ensurePcpDashboardDateValue())}`;

  const snapshot = state.pcpDashboardSnapshot;
  const cnc = snapshot && snapshot.charts ? snapshot.charts.cnc : null;
  const machines = Array.isArray(cnc) ? cnc : [];
  const rncUsinagem = snapshot && snapshot.charts ? snapshot.charts.rncUsinagem : null;
  const rncMachines = (rncUsinagem && Array.isArray(rncUsinagem.machines)) ? rncUsinagem.machines : [];
  const rncTotals = (rncUsinagem && rncUsinagem.totals && typeof rncUsinagem.totals === 'object') ? rncUsinagem.totals : {};
  if (!machines.length) {
    elements.pcpCncEmpty.style.display = 'block';
    elements.pcpCncEmpty.textContent = 'Sem dados de usinagem por máquina para a data selecionada.';
    return;
  }
  elements.pcpCncEmpty.style.display = 'none';
  elements.pcpCncOverview.appendChild(createDashboardPdfSectionHeader('USINAGEM MÁQUINAS CNC', ensurePcpDashboardDateValue()));

  const findMachine = (name) => machines.find((m) => m.machine === name) || {};
  const getMachineColor = (item) => CNC_MACHINE_COLOR_MAP[item.machine] || '#64748b';
  const kpiRow = document.createElement('div');
  kpiRow.className = 'pcp-dashboard-kpi-row cols-5';
  kpiRow.appendChild(createDashboardKpiCard('Efic Fagor 1', formatPercent(findMachine('FAGOR 1').efficiencyOps || 0)));
  kpiRow.appendChild(createDashboardKpiCard('Efic Fagor 2', formatPercent(findMachine('FAGOR 2').efficiencyOps || 0)));
  kpiRow.appendChild(createDashboardKpiCard('Efic Mcs 1', formatPercent(findMachine('MCS 1').efficiencyOps || 0)));
  kpiRow.appendChild(createDashboardKpiCard('Efic Mcs 2', formatPercent(findMachine('MCS 2').efficiencyOps || 0)));
  kpiRow.appendChild(createDashboardKpiCard('Efic Mcs 3', formatPercent(findMachine('MCS 3').efficiencyOps || 0)));
  elements.pcpCncOverview.appendChild(kpiRow);

  const gapPecas = machines.reduce(
    (acc, m) => acc + Math.max(0, Number(m.piecesPlanned || 0) - Number(m.piecesDone || 0)),
    0
  );
  const gapOps = machines.reduce(
    (acc, m) => acc + Math.max(0, Number(m.opsPlanned || 0) - Number(m.opsDone || 0)),
    0
  );

  const kpiRowExec = document.createElement('div');
  kpiRowExec.className = 'pcp-dashboard-kpi-row cols-4';
  kpiRowExec.appendChild(createDashboardKpiCard('Gap de Peças (Planejadas - Produzidas)', formatInteger(gapPecas)));
  kpiRowExec.appendChild(createDashboardKpiCard('Gap de OPs (Planejadas - Produzidas)', formatInteger(gapOps)));
  kpiRowExec.appendChild(createDashboardKpiCard('Peças Refugadas (RNC/USINAGEM)', formatInteger(rncTotals.piecesRefugadas || 0)));
  kpiRowExec.appendChild(createDashboardKpiCard('Tempo Perdido (RNC/USINAGEM)', formatDurationFromSeconds(rncTotals.segsPerdidos || 0)));
  elements.pcpCncOverview.appendChild(kpiRowExec);

  const mainRow = document.createElement('div');
  mainRow.className = 'pcp-cnc-main-row';
  mainRow.appendChild(createCncVerticalComparisonChart({
    title: 'Quantidade de peça feita por máquina (Produzidas X Planejadas)',
    data: machines,
    doneKey: 'piecesDone',
    plannedKey: 'piecesPlanned',
    valueFormatter: (v) => formatInteger(v),
    colorResolver: getMachineColor
  }));
  mainRow.appendChild(createCncVerticalComparisonChart({
    title: 'Quantidade de OPS por máquina (Produzidas X Planejadas)',
    data: machines,
    doneKey: 'opsDone',
    plannedKey: 'opsPlanned',
    valueFormatter: (v) => formatInteger(v),
    colorResolver: getMachineColor
  }));
  elements.pcpCncCharts.appendChild(mainRow);

  if (!rncMachines.length) {
    elements.pcpCncCharts.appendChild(createDashboardPlaceholderChart('RNC por máquina (USINAGEM)', 'Sem dados de RNC para a data selecionada.'));
    elements.pcpCncCharts.appendChild(createDashboardPlaceholderChart('Tempo perdido por máquina (USINAGEM)', 'Sem dados de RNC para a data selecionada.'));
  } else {
    elements.pcpCncCharts.appendChild(createDashboardBarChart({
      title: 'RNC por máquina (USINAGEM)',
      data: rncMachines.map((item) => ({ name: item.machine, value: item.rnc || 0 })),
      colorResolver: () => '#ef4444',
      valueFormatter: (v) => formatInteger(v)
    }));

    elements.pcpCncCharts.appendChild(createDashboardBarChart({
      title: 'Tempo perdido por máquina (USINAGEM)',
      data: rncMachines.map((item) => ({ name: item.machine, value: item.segs || 0 })),
      colorResolver: () => '#f59e0b',
      valueFormatter: (v) => formatDurationFromSeconds(v)
    }));
  }

  const bottomRow = document.createElement('div');
  bottomRow.className = 'pcp-cnc-bottom-row';
  bottomRow.appendChild(createCncContributionChart(machines));
  bottomRow.appendChild(rncMachines.length
    ? createDashboardBarChart({
      title: 'Custo apurado por máquina (RNC/USINAGEM)',
      data: rncMachines.map((item) => ({ name: item.machine, value: item.cost || 0 })),
      colorResolver: () => '#22c55e',
      valueFormatter: (v) => formatCurrencyBr(v)
    })
    : createDashboardPlaceholderChart('Custo apurado por máquina (RNC/USINAGEM)', 'Sem dados de RNC para a data selecionada.'));
  elements.pcpCncCharts.appendChild(bottomRow);
}

async function loadPcpMoldagemSnapshot() {
  const date = ensurePcpMoldagemDateValue();
  state.pcpMoldagemSnapshot = await window.api.loadPcpMoldagem({ date });
}

async function loadPcpDashboardSnapshot() {
  const date = ensurePcpDashboardDateValue();
  state.pcpDashboardSnapshot = await window.api.loadPcpDashboard({ date });
}

async function openPcpDashboard() {
  if (!state.pcpDashboardSnapshot) {
    await loadPcpDashboardSnapshot();
  }
  if (!state.pcpMoldagemSnapshot) {
    await loadPcpMoldagemSnapshot();
  }
  setPcpView('dashboard');
  renderPcpDashboard();
}

async function exportPcpDashboardPdf() {
  const date = ensurePcpDashboardDateValue();
  const previousTheme = elements.body.getAttribute('data-theme') || 'light';
  const previousPdfHeaderHidden = elements.pcpDashboardPdfHeader ? elements.pcpDashboardPdfHeader.hidden : true;
  const previousPdfTitle = elements.pcpDashboardPdfTitle ? elements.pcpDashboardPdfTitle.textContent : '';

  elements.body.classList.add('pdf-dashboard-export');
  elements.body.setAttribute('data-theme', 'light');
  if (elements.pcpDashboardPdfTitle) {
    elements.pcpDashboardPdfTitle.textContent = `VISÃO GERAL - DASHBOARD - ${formatWeekdayAndDatePtBr(date)}`;
  }
  if (elements.pcpDashboardPdfHeader) {
    elements.pcpDashboardPdfHeader.hidden = false;
  }

  try {
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const result = await window.api.exportPcpDashboardPdf({
      suggestedName: `pcp-dashboard-${date}.pdf`,
      exportWidth: 1122,
      exportHeight: 794
    });
    if (!result || result.canceled) {
      return null;
    }
    return result.filePath;
  } finally {
    elements.body.classList.remove('pdf-dashboard-export');
    elements.body.setAttribute('data-theme', previousTheme);
    if (elements.pcpDashboardPdfHeader) {
      elements.pcpDashboardPdfHeader.hidden = previousPdfHeaderHidden;
    }
    if (elements.pcpDashboardPdfTitle) {
      elements.pcpDashboardPdfTitle.textContent = previousPdfTitle || 'VISÃO GERAL - DASHBOARD';
    }
    await new Promise((resolve) => requestAnimationFrame(resolve));
  }
}

function getCaptureRectForElement(element) {
  const rect = element.getBoundingClientRect();
  const fullWidth = Math.max(
    Math.ceil(rect.width),
    Math.ceil(element.scrollWidth || 0),
    Math.ceil(element.offsetWidth || 0)
  );
  const fullHeight = Math.max(
    Math.ceil(rect.height),
    Math.ceil(element.scrollHeight || 0),
    Math.ceil(element.offsetHeight || 0)
  );
  return {
    x: Math.max(0, Math.floor(rect.left + window.scrollX)),
    y: Math.max(0, Math.floor(rect.top + window.scrollY)),
    width: Math.max(1, fullWidth),
    height: Math.max(1, fullHeight)
  };
}

async function exportPcpEfficiencyImage() {
  if (!elements.pcpEffExportArea) {
    throw new Error('Area de exportacao nao encontrada.');
  }
  if (!state.pcpEfficiencySnapshot) {
    throw new Error('Carregue os dados da eficiencia antes de exportar.');
  }

  renderPcpEfficiency();
  elements.pcpEffExportArea.scrollIntoView({ block: 'start', behavior: 'auto' });
  await new Promise((resolve) => requestAnimationFrame(() => resolve()));
  const rect = getCaptureRectForElement(elements.pcpEffExportArea);
  const date = ensurePcpEffDateValue();
  const desiredExportWidth = Math.max(rect.x + rect.width + 420, 1620);
  const desiredExportHeight = Math.max(rect.y + rect.height + 380, 1160);
  const result = await window.api.exportPcpEfficiencyImage({
    rect,
    selector: '#pcp-eff-export-area',
    autoFit: true,
    targetWidth: desiredExportWidth,
    targetHeight: desiredExportHeight,
    suggestedName: `pcp-eficiencia-${date}.png`
  });

  if (!result || result.canceled) {
    return null;
  }
  return result.filePath;
}

async function handleAppLoginConfirm() {
  const username = String(elements.appLoginUsername?.value || '').trim();
  const password = String(elements.appLoginPassword?.value || '').trim();
  if (!username || !password) {
    if (elements.appLoginError) {
      elements.appLoginError.textContent = 'Informe usuário e senha.';
    }
    return;
  }

  const result = await window.api.verifyLogin({ username, password, requireFiscal: false });
  if (!result || !result.ok) {
    if (elements.appLoginError) {
      elements.appLoginError.textContent = (result && result.error) || 'Credenciais inválidas.';
    }
    const errText = String((result && result.error) || '').toLowerCase();
    if (errText.includes('banco pedidos nao encontrado')) {
      showToast('Banco Pedidos não encontrado. Configure o Caminho Fiscal (R:) em Configuração.', true);
    }
    elements.appLoginPassword?.focus?.();
    elements.appLoginPassword?.select?.();
    return;
  }

  state.authUser = {
    username: result.username,
    permissions: result.permissions || {}
  };
  state.fiscalUser = result.username;
  applyPermissionsToTabs();

  try {
    state.settings = await window.api.saveSettings({ lastLoginUsername: result.username });
  } catch (error) {
    // Ignore persistence errors for last login name.
  }

  closeAppLoginModal();
  switchTab('menu');
  showToast(`Bem-vindo, ${result.username}.`);
}

function handleAppLoginOpenSettings() {
  closeAppLoginModal();
  switchTab('settings');
  elements.btnSelectFiscalRoot?.focus?.();
}

function requestPcpAccess() {
  if (!canAccessPcp()) {
    showToast('Usuário sem permissão para acessar o PCP.', true);
    return;
  }
  setPcpView('menu');
  switchTab('pcp');
}

async function requestFiscalAccess() {
  if (!canAccessFiscal()) {
    showToast('Usuário sem permissão para acessar o FISCAL.', true);
    return;
  }

  switchTab('fiscal');
  try {
    await loadFiscalNfs();
  } catch (error) {
    showToast(error.message || 'Erro ao carregar NFs do banco.', true);
  }
  state.fiscalSelectionNf = null;
  state.fiscalNfItems = [];
  state.fiscalView = 'menu';
  if (elements.fiscalOpenCompras) {
    elements.fiscalOpenCompras.hidden = !canFiscalComprasView();
  }
  renderFiscalScreen();
}

function getFiscalComprasPermissionValue(key) {
  const perms = getAuthPermissions();
  const hasKey = perms && Object.prototype.hasOwnProperty.call(perms, key);
  const hasFiscal = !!getPermissionValue(perms, 'fiscal');
  if (hasKey) {
    return !!perms[key];
  }
  return hasFiscal;
}

function canFiscalComprasView() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_view');
}

function canFiscalComprasCreate() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_create');
}

function canFiscalComprasEdit() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_edit');
}

function canFiscalComprasDelete() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_delete');
}

function canFiscalComprasPdf() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_pdf');
}

function canFiscalComprasApprove() {
  return canAccessFiscal() && getFiscalComprasPermissionValue('compras_approve');
}

function parseMoneyLike(value) {
  const raw = String(value || '').trim();
  if (!raw) return 0;
  const cleaned = raw.replace(/[^\d,.\-]/g, '');
  const lastComma = cleaned.lastIndexOf(',');
  const lastDot = cleaned.lastIndexOf('.');
  const decimalSep = lastComma > lastDot ? ',' : '.';
  let normalized = cleaned;
  if (lastComma >= 0 || lastDot >= 0) {
    const thousandsSep = decimalSep === ',' ? '.' : ',';
    normalized = normalized.split(thousandsSep).join('');
    if (decimalSep === ',') {
      normalized = normalized.replace(',', '.');
    }
  }
  const num = Number(normalized);
  return Number.isFinite(num) ? num : 0;
}

function formatStatusCompraLabel(value) {
  const key = String(value || '').trim().toUpperCase();
  if (key === 'PENDENTE') return 'Pendente';
  if (key === 'APROVADA') return 'Aprovada';
  if (key === 'CANCELADA') return 'Cancelada';
  return key || '-';
}

function createFiscalCompraItemSeed() {
  return {
    quantidade: '',
    unidade: '',
    descricao: '',
    fornecedor_escolhido: '',
    cotacoes: [
      { fornecedor_nome: '', tipo_pagamento: '', valor_unitario: '', prazo_entrega: '' },
      { fornecedor_nome: '', tipo_pagamento: '', valor_unitario: '', prazo_entrega: '' },
      { fornecedor_nome: '', tipo_pagamento: '', valor_unitario: '', prazo_entrega: '' }
    ]
  };
}

function setFiscalComprasTab(tab) {
  state.fiscalComprasTab = tab === 'consulta' ? 'consulta' : 'form';
  if (elements.fiscalComprasViewForm) {
    elements.fiscalComprasViewForm.hidden = state.fiscalComprasTab !== 'form';
  }
  if (elements.fiscalComprasViewConsulta) {
    elements.fiscalComprasViewConsulta.hidden = state.fiscalComprasTab !== 'consulta';
  }
  elements.fiscalComprasTabForm?.classList.toggle('active', state.fiscalComprasTab === 'form');
  elements.fiscalComprasTabConsulta?.classList.toggle('active', state.fiscalComprasTab === 'consulta');
}

function getFornecedorPagamentoPadrao(nome) {
  const safe = normalizeForFilter(nome);
  const found = (state.fiscalComprasFornecedores || []).find((f) => normalizeForFilter(f.nome) === safe);
  return found ? String(found.tipo_pagamento_padrao || '').trim() : '';
}

function renderFiscalComprasDatalists() {
  if (elements.fiscalComprasFornecedorList) {
    elements.fiscalComprasFornecedorList.innerHTML = '';
    (state.fiscalComprasFornecedores || []).forEach((f) => {
      const option = document.createElement('option');
      option.value = String(f.nome || '').trim();
      elements.fiscalComprasFornecedorList.appendChild(option);
    });
  }
  if (elements.fiscalComprasPagamentoList) {
    const unique = new Set();
    elements.fiscalComprasPagamentoList.innerHTML = '';
    (state.fiscalComprasFornecedores || []).forEach((f) => {
      const value = String(f.tipo_pagamento_padrao || '').trim();
      if (!value || unique.has(normalizeForFilter(value))) {
        return;
      }
      unique.add(normalizeForFilter(value));
      const option = document.createElement('option');
      option.value = value;
      elements.fiscalComprasPagamentoList.appendChild(option);
    });
  }
}

function recalcFiscalComprasItemCard(card) {
  const qtyInput = card.querySelector('[data-item-qty]');
  const qty = parseMoneyLike(qtyInput ? qtyInput.value : 0);
  const quoteRows = Array.from(card.querySelectorAll('[data-cot-row]'));
  let minUnit = Number.POSITIVE_INFINITY;
  quoteRows.forEach((row) => {
    const unitInput = row.querySelector('[data-cot-unit]');
    const subtotalEl = row.querySelector('[data-cot-subtotal]');
    const unit = parseMoneyLike(unitInput ? unitInput.value : 0);
    const subtotal = qty > 0 && unit > 0 ? qty * unit : 0;
    if (subtotalEl) {
      subtotalEl.textContent = subtotal > 0 ? formatCurrencyBr(subtotal) : '-';
    }
    row.classList.remove('best-price');
    if (unit > 0) {
      minUnit = Math.min(minUnit, unit);
    }
  });
  if (Number.isFinite(minUnit)) {
    quoteRows.forEach((row) => {
      const unitInput = row.querySelector('[data-cot-unit]');
      const unit = parseMoneyLike(unitInput ? unitInput.value : 0);
      if (unit > 0 && Math.abs(unit - minUnit) < 0.000001) {
        row.classList.add('best-price');
      }
    });
  }
}

function recalcFiscalComprasTotalForm() {
  const items = readFiscalComprasItemsFromForm();
  const total = items.reduce((sum, item) => {
    const escolha = String(item.fornecedor_escolhido || '').trim();
    const chosen = item.cotacoes.find((c) => String(c.fornecedor_nome || '').trim() === escolha) || item.cotacoes[0];
    return sum + Number(chosen && chosen.subtotal ? chosen.subtotal : 0);
  }, 0);
  if (elements.fiscalComprasTotalEstimado) {
    elements.fiscalComprasTotalEstimado.value = formatCurrencyBr(total);
  }
}

function renderFiscalComprasItems(itemsSeed) {
  if (!elements.fiscalComprasItemsWrap) {
    return;
  }
  const items = Array.isArray(itemsSeed) && itemsSeed.length ? itemsSeed : [createFiscalCompraItemSeed()];
  elements.fiscalComprasItemsWrap.innerHTML = '';

  items.forEach((item, index) => {
    const card = document.createElement('div');
    card.className = 'fiscal-compra-item-card';
    card.setAttribute('data-item-card', String(index));

    const head = document.createElement('div');
    head.className = 'inline';
    head.style.justifyContent = 'space-between';
    head.style.alignItems = 'center';
    const title = document.createElement('h4');
    title.className = 'fiscal-compra-item-title';
    title.textContent = `Item ${index + 1}`;
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn secondary';
    removeBtn.textContent = 'Remover';
    removeBtn.disabled = items.length <= 1;
    removeBtn.addEventListener('click', () => {
      const next = readFiscalComprasItemsFromForm();
      next.splice(index, 1);
      renderFiscalComprasItems(next);
      recalcFiscalComprasTotalForm();
    });
    head.appendChild(title);
    head.appendChild(removeBtn);
    card.appendChild(head);

    const core = document.createElement('div');
    core.className = 'inline';
    core.innerHTML = `
      <div class="field" style="flex:1;">
        <label>Quantidade</label>
        <input type="number" min="0" step="0.001" data-item-qty value="${String(item.quantidade || '')}" />
      </div>
      <div class="field" style="flex:1;">
        <label>Unidade</label>
        <input type="text" data-item-unidade value="${String(item.unidade || '')}" />
      </div>
      <div class="field" style="flex:3;">
        <label>Descrição</label>
        <input type="text" data-item-descricao value="${String(item.descricao || '')}" />
      </div>
    `;
    card.appendChild(core);

    const cotGrid = document.createElement('div');
    cotGrid.className = 'fiscal-compra-cotacao-grid';
    const cotacoes = Array.isArray(item.cotacoes) ? item.cotacoes.slice(0, 3) : [];
    while (cotacoes.length < 3) {
      cotacoes.push({ fornecedor_nome: '', tipo_pagamento: '', valor_unitario: '', prazo_entrega: '' });
    }
    cotacoes.forEach((cot, cotIndex) => {
      const row = document.createElement('div');
      row.className = 'fiscal-compra-cotacao-row';
      row.setAttribute('data-cot-row', String(cotIndex));
      row.innerHTML = `
        <div class="field">
          <label>Fornecedor ${cotIndex + 1}</label>
          <input type="text" list="fiscal-compras-fornecedor-list" data-cot-fornecedor value="${String(cot.fornecedor_nome || '')}" />
        </div>
        <div class="field">
          <label>Pagamento</label>
          <input type="text" list="fiscal-compras-pagamento-list" data-cot-pagamento value="${String(cot.tipo_pagamento || '')}" />
        </div>
        <div class="field">
          <label>Valor unit.</label>
          <input type="number" min="0" step="0.01" data-cot-unit value="${String(cot.valor_unitario || '')}" />
        </div>
        <div class="field">
          <label>Prazo</label>
          <input type="text" data-cot-prazo value="${String(cot.prazo_entrega || '')}" />
        </div>
        <div class="field">
          <label>Subtotal</label>
          <div class="fiscal-cot-subtotal-value" data-cot-subtotal>-</div>
        </div>
        <div class="field fiscal-cot-choose-field">
          <label>Escolha</label>
          <input class="fiscal-cot-choose-input" type="radio" name="item-escolha-${index}" data-cot-escolha value="${cotIndex}" ${String(item.fornecedor_escolhido || '') === String(cot.fornecedor_nome || '') ? 'checked' : ''} />
        </div>
      `;
      cotGrid.appendChild(row);
    });
    card.appendChild(cotGrid);

    card.querySelectorAll('input').forEach((input) => {
      input.addEventListener('input', () => {
        if (input.hasAttribute('data-cot-fornecedor')) {
          const fornecedor = String(input.value || '').trim();
          const pagamentoInput = input.closest('[data-cot-row]')?.querySelector('[data-cot-pagamento]');
          if (pagamentoInput && !String(pagamentoInput.value || '').trim()) {
            pagamentoInput.value = getFornecedorPagamentoPadrao(fornecedor);
          }
        }
        recalcFiscalComprasItemCard(card);
        recalcFiscalComprasTotalForm();
      });
      input.addEventListener('change', () => {
        recalcFiscalComprasItemCard(card);
        recalcFiscalComprasTotalForm();
      });
    });

    const radios = Array.from(card.querySelectorAll('[data-cot-escolha]'));
    if (!radios.some((r) => r.checked) && radios.length) {
      radios[0].checked = true;
    }

    recalcFiscalComprasItemCard(card);
    elements.fiscalComprasItemsWrap.appendChild(card);
  });
}

function readFiscalComprasItemsFromForm() {
  const cards = Array.from(elements.fiscalComprasItemsWrap?.querySelectorAll('[data-item-card]') || []);
  return cards.map((card) => {
    const quantidade = parseMoneyLike(card.querySelector('[data-item-qty]')?.value);
    const unidade = String(card.querySelector('[data-item-unidade]')?.value || '').trim();
    const descricao = String(card.querySelector('[data-item-descricao]')?.value || '').trim();
    const cotRows = Array.from(card.querySelectorAll('[data-cot-row]'));
    const cotacoes = cotRows.map((row) => {
      const fornecedor_nome = String(row.querySelector('[data-cot-fornecedor]')?.value || '').trim();
      const tipo_pagamento = String(row.querySelector('[data-cot-pagamento]')?.value || '').trim();
      const valor_unitario = parseMoneyLike(row.querySelector('[data-cot-unit]')?.value);
      const prazo_entrega = String(row.querySelector('[data-cot-prazo]')?.value || '').trim();
      return {
        fornecedor_nome,
        tipo_pagamento,
        valor_unitario,
        prazo_entrega,
        subtotal: valor_unitario > 0 && quantidade > 0 ? quantidade * valor_unitario : 0
      };
    });
    const selectedIdx = cotRows.findIndex((row) => !!row.querySelector('[data-cot-escolha]')?.checked);
    const selected = selectedIdx >= 0 ? cotacoes[selectedIdx] : cotacoes[0];
    return {
      quantidade,
      unidade,
      descricao,
      fornecedor_escolhido: selected ? selected.fornecedor_nome : '',
      cotacoes: cotacoes.filter((c) => c.fornecedor_nome || c.tipo_pagamento || c.valor_unitario > 0 || c.prazo_entrega)
    };
  });
}

function resetFiscalComprasForm() {
  if (elements.fiscalComprasError) {
    elements.fiscalComprasError.textContent = '';
  }
  if (elements.fiscalComprasId) elements.fiscalComprasId.value = '';
  if (elements.fiscalComprasNumero) elements.fiscalComprasNumero.value = '';
  if (elements.fiscalComprasDataSolic) elements.fiscalComprasDataSolic.value = getTodayYmdLocal();
  if (elements.fiscalComprasSolicitadoPor) elements.fiscalComprasSolicitadoPor.value = getAuthUsername();
  if (elements.fiscalComprasSetor) elements.fiscalComprasSetor.value = '';
  if (elements.fiscalComprasDestinado) elements.fiscalComprasDestinado.value = '';
  if (elements.fiscalComprasStatus) elements.fiscalComprasStatus.value = 'PENDENTE';
  if (elements.fiscalComprasObservacoes) elements.fiscalComprasObservacoes.value = '';
  renderFiscalComprasItems([createFiscalCompraItemSeed()]);
  recalcFiscalComprasTotalForm();
}

function fillFiscalComprasForm(row) {
  if (!row || typeof row !== 'object') {
    resetFiscalComprasForm();
    return;
  }
  if (elements.fiscalComprasError) elements.fiscalComprasError.textContent = '';
  if (elements.fiscalComprasId) elements.fiscalComprasId.value = String(row.id || '').trim();
  if (elements.fiscalComprasNumero) elements.fiscalComprasNumero.value = String(row.numero_requisicao || '').trim();
  if (elements.fiscalComprasDataSolic) elements.fiscalComprasDataSolic.value = String(row.data_solicitacao || '').trim();
  if (elements.fiscalComprasSolicitadoPor) elements.fiscalComprasSolicitadoPor.value = String(row.solicitado_por || '').trim();
  if (elements.fiscalComprasSetor) elements.fiscalComprasSetor.value = String(row.setor || '').trim();
  if (elements.fiscalComprasDestinado) elements.fiscalComprasDestinado.value = String(row.destinado_a || '').trim();
  if (elements.fiscalComprasStatus) elements.fiscalComprasStatus.value = String(row.status || 'PENDENTE').trim();
  if (elements.fiscalComprasObservacoes) elements.fiscalComprasObservacoes.value = String(row.observacoes || '').trim();
  renderFiscalComprasItems(Array.isArray(row.itens) && row.itens.length ? row.itens : [createFiscalCompraItemSeed()]);
  recalcFiscalComprasTotalForm();
}

async function loadFiscalComprasContext() {
  const username = getAuthUsername();
  const [ctxResult, fornResult] = await Promise.all([
    window.api.fiscalComprasContext({ username }),
    window.api.fiscalComprasFornecedoresList({ username })
  ]);
  state.fiscalComprasCan = (ctxResult && ctxResult.permissions) || {};
  state.fiscalComprasFornecedores = (fornResult && fornResult.rows) || [];
  renderFiscalComprasDatalists();
}

function applyFiscalComprasPermissionsUi() {
  if (elements.fiscalComprasNew) {
    elements.fiscalComprasNew.hidden = !canFiscalComprasCreate();
  }
  if (elements.fiscalComprasAddItem) {
    elements.fiscalComprasAddItem.disabled = !canFiscalComprasCreate() && !canFiscalComprasEdit();
  }
  if (elements.fiscalComprasSave) {
    elements.fiscalComprasSave.disabled = !canFiscalComprasCreate() && !canFiscalComprasEdit();
  }
  if (elements.fiscalComprasSavePdf) {
    elements.fiscalComprasSavePdf.disabled = (!canFiscalComprasCreate() && !canFiscalComprasEdit()) || !canFiscalComprasPdf();
  }
}

function renderFiscalComprasTable() {
  if (!elements.fiscalComprasTableHead || !elements.fiscalComprasTableBody) {
    return;
  }
  setTableHeaders(elements.fiscalComprasTableHead, ['Nº Req.', 'Data', 'Solicitante', 'Setor', 'Status', 'Total', 'Ações']);
  elements.fiscalComprasTableBody.innerHTML = '';
  const rows = Array.isArray(state.fiscalComprasRows) ? state.fiscalComprasRows : [];
  if (elements.fiscalComprasConsultaEmpty) {
    elements.fiscalComprasConsultaEmpty.hidden = rows.length > 0;
  }
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    const values = [
      String(row.numero_requisicao || '').trim(),
      formatYmdToBr(String(row.data_solicitacao || '').trim()),
      String(row.solicitado_por || '').trim(),
      String(row.setor || '').trim(),
      formatStatusCompraLabel(row.status),
      formatCurrencyBr(Number(row.total_geral_estimado || 0))
    ];
    values.forEach((value, idx) => {
      const td = document.createElement('td');
      if (idx === 4) {
        const statusText = String(value || '').trim();
        const statusBadge = document.createElement('span');
        statusBadge.className = `fiscal-compras-status-badge status-${String(row.status || '').trim().toLowerCase()}`;
        statusBadge.textContent = statusText || '-';
        td.appendChild(statusBadge);
        if (String(row.status || '').trim().toUpperCase() === 'APROVADA' && String(row.aprovacao_obs || '').trim()) {
          const obsBadge = document.createElement('span');
          obsBadge.className = 'fiscal-compras-aprov-obs';
          const fullObs = String(row.aprovacao_obs || '').trim();
          obsBadge.textContent = fullObs;
          obsBadge.title = fullObs;
          td.appendChild(obsBadge);
        }
      } else {
        td.textContent = value;
      }
      tr.appendChild(td);
    });

    const actionsTd = document.createElement('td');
    actionsTd.className = 'fiscal-compras-actions-cell';
    const actionsWrap = document.createElement('div');
    actionsWrap.className = 'fiscal-compras-actions-wrap';

    const topRow = document.createElement('div');
    topRow.className = 'fiscal-compras-actions-top';
    const canApproveObs = canFiscalComprasApprove();
    const canChangeStatus = canFiscalComprasEdit() || canFiscalComprasApprove();
    const statusSel = document.createElement('select');
    statusSel.className = 'fiscal-compras-actions-status';
    const originalStatus = String(row.status || '').trim();
    ['PENDENTE', 'APROVADA', 'CANCELADA'].forEach((status) => {
      const opt = document.createElement('option');
      opt.value = status;
      opt.textContent = formatStatusCompraLabel(status);
      if (String(row.status || '').trim() === status) {
        opt.selected = true;
      }
      statusSel.appendChild(opt);
    });
    statusSel.disabled = !canChangeStatus;
    topRow.appendChild(statusSel);

    const updateBtn = document.createElement('button');
    updateBtn.type = 'button';
    updateBtn.className = 'btn secondary fiscal-compras-actions-update';
    updateBtn.textContent = 'Atualizar';
    updateBtn.disabled = !canChangeStatus;
    updateBtn.addEventListener('click', async () => {
      const prevObs = String((obsInput && obsInput.value) || String(row.aprovacao_obs || '').trim());
      try {
        const payload = {
          username: getAuthUsername(),
          id: row.id,
          status: statusSel.value
        };
        if (canApproveObs) {
          payload.aprovacao_obs = String(obsInput.value || '').trim();
        }
        await window.api.fiscalComprasUpdateStatus(payload);
        await loadFiscalComprasRows();
        showToast('Status atualizado.');
      } catch (error) {
        statusSel.value = originalStatus;
        if (canApproveObs) {
          obsInput.value = prevObs;
        }
        showToast(error.message || 'Erro ao atualizar status.', true);
      }
    });
    topRow.appendChild(updateBtn);
    actionsWrap.appendChild(topRow);

    const obsRow = document.createElement('div');
    obsRow.className = 'fiscal-compras-actions-obs-row';
    const obsLabel = document.createElement('span');
    obsLabel.className = 'fiscal-compras-actions-obs-label';
    obsLabel.textContent = 'Obs aprovação';
    obsRow.appendChild(obsLabel);
    const obsInput = document.createElement('input');
    obsInput.type = 'text';
    obsInput.className = 'fiscal-compras-actions-obs';
    obsInput.placeholder = canApproveObs ? 'Ex: chega 25/02' : 'Sem observação';
    obsInput.value = String(row.aprovacao_obs || '').trim();
    obsInput.title = String(row.aprovacao_obs || '').trim();
    if (!canApproveObs) {
      obsInput.readOnly = true;
      obsInput.classList.add('readonly');
    }
    obsRow.appendChild(obsInput);
    actionsWrap.appendChild(obsRow);

    const obsHint = document.createElement('p');
    obsHint.className = 'fiscal-compras-actions-hint';
    obsHint.textContent = canApproveObs
      ? 'Somente aprovador edita a observação.'
      : 'Somente aprovador pode editar a observação.';
    actionsWrap.appendChild(obsHint);

    const linksRow = document.createElement('div');
    linksRow.className = 'fiscal-compras-actions-links';

    const editBtn = createPcpActionButton('Editar', () => {
      fillFiscalComprasForm(row);
      setFiscalComprasTab('form');
    });
    editBtn.classList.add('fiscal-compras-link-action');
    editBtn.disabled = !canFiscalComprasEdit();
    linksRow.appendChild(editBtn);

    const pdfBtn = createPcpActionButton('PDF', async () => {
      try {
        await window.api.fiscalComprasPdfRegen({ username: getAuthUsername(), id: row.id });
        showToast('PDF gerado.');
      } catch (error) {
        showToast(error.message || 'Erro ao gerar PDF.', true);
      }
    });
    pdfBtn.classList.add('fiscal-compras-link-action');
    pdfBtn.disabled = !canFiscalComprasPdf();
    linksRow.appendChild(pdfBtn);

    const openBtn = createPcpActionButton('Abrir PDF', async () => {
      try {
        await window.api.fiscalComprasPdfOpen({ username: getAuthUsername(), id: row.id });
      } catch (error) {
        showToast(error.message || 'Erro ao abrir PDF.', true);
      }
    });
    openBtn.classList.add('fiscal-compras-link-action');
    openBtn.disabled = !canFiscalComprasPdf();
    linksRow.appendChild(openBtn);

    const histBtn = createPcpActionButton('Histórico', () => {
      const entries = Array.isArray(row.historico_itens) ? row.historico_itens : [];
      if (!entries.length) {
        showToast('Sem histórico de compra para os itens desta solicitação.');
        return;
      }
      const text = entries.map((h) => (
        `${h.descricao}: ${h.ultimo_fornecedor || '-'} / ${h.ultimo_valor_pago ? formatCurrencyBr(h.ultimo_valor_pago) : '-'} / ${formatYmdToBr(h.data_ultima_compra || '')}`
      )).join('\n');
      window.alert(text);
    });
    histBtn.classList.add('fiscal-compras-link-action');
    linksRow.appendChild(histBtn);

    const delBtn = createPcpActionButton('Excluir', async () => {
      const ok = window.confirm('Excluir esta solicitação de compra?');
      if (!ok) return;
      try {
        await window.api.fiscalComprasDelete({ username: getAuthUsername(), id: row.id });
        await loadFiscalComprasRows();
        showToast('Solicitação excluída.');
      } catch (error) {
        showToast(error.message || 'Erro ao excluir solicitação.', true);
      }
    });
    delBtn.classList.add('fiscal-compras-link-action');
    delBtn.disabled = !canFiscalComprasDelete();
    linksRow.appendChild(delBtn);

    actionsWrap.appendChild(linksRow);
    actionsTd.appendChild(actionsWrap);

    tr.appendChild(actionsTd);
    elements.fiscalComprasTableBody.appendChild(tr);
  });
}

async function loadFiscalComprasRows() {
  const result = await window.api.fiscalComprasList({
    username: getAuthUsername(),
    filters: state.fiscalComprasFilters || {}
  });
  state.fiscalComprasRows = (result && result.rows) || [];
  renderFiscalComprasTable();
}

async function openFiscalComprasView() {
  if (!canFiscalComprasView()) {
    showToast('Usuário sem permissão para visualizar Solicitação de Compras.', true);
    return;
  }
  state.fiscalView = 'compras';
  renderFiscalScreen();
  await loadFiscalComprasContext();
  applyFiscalComprasPermissionsUi();
  await loadFiscalComprasRows();
  resetFiscalComprasForm();
  setFiscalComprasTab('consulta');
}

function canEditRh() {
  const perms = getAuthPermissions();
  return !!(perms && Object.prototype.hasOwnProperty.call(perms, 'rh_edit') ? perms.rh_edit : false);
}

function setRhView(next) {
  state.rhView = next;
  if (elements.rhMenuView) {
    elements.rhMenuView.hidden = next !== 'menu';
  }
  if (elements.rhHeView) {
    elements.rhHeView.hidden = next !== 'he';
  }
  if (elements.rhSolicHeView) {
    elements.rhSolicHeView.hidden = next !== 'solic_he';
  }
}

function setRhHeSubView(next) {
  state.rhHeSubView = next;
  const views = {
    dashboard: elements.rhViewDashboard,
    atrasos: elements.rhViewAtrasos,
    registros: elements.rhViewRegistros,
    colaboradores: elements.rhViewColaboradores,
    setores: elements.rhViewSetores,
    config: elements.rhViewConfig,
    audit: elements.rhViewAudit
  };
  const tabs = {
    dashboard: elements.rhSubtabDashboard,
    atrasos: elements.rhSubtabAtrasos,
    registros: elements.rhSubtabRegistros,
    colaboradores: elements.rhSubtabColaboradores,
    setores: elements.rhSubtabSetores,
    config: elements.rhSubtabConfig,
    audit: elements.rhSubtabAudit
  };
  Object.entries(views).forEach(([key, el]) => {
    if (el) {
      el.hidden = key !== next;
    }
  });
  Object.entries(tabs).forEach(([key, el]) => {
    el?.classList.toggle('active', key === next);
  });
}

function requestRhAccess() {
  if (!canAccessArea('rh')) {
    showToast('Usuário sem permissão para acessar RH.', true);
    return;
  }
  setRhView('menu');
  setRhHeSubView('dashboard');
  switchTab('rh');
}

function getAuthUsername() {
  return String(state.authUser?.username || '').trim();
}

function getCurrentMonthValue() {
  const now = new Date();
  const yyyy = String(now.getFullYear()).padStart(4, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  return `${yyyy}-${mm}`;
}

function ensureRhMonthValue() {
  if (!elements.rhFilterMonth) {
    return getCurrentMonthValue();
  }
  if (elements.rhFilterMonth.value) {
    return elements.rhFilterMonth.value;
  }
  const value = getCurrentMonthValue();
  elements.rhFilterMonth.value = value;
  return value;
}

function getRhUiFilters() {
  return {
    month: ensureRhMonthValue(),
    setorId: String(elements.rhFilterSector?.value || '').trim(),
    colaboradorId: String(elements.rhFilterCollab?.value || '').trim(),
    tipoHora: String(elements.rhFilterType?.value || '').trim()
  };
}

function setRhSelectOptions(selectEl, options, { includeAll = true, allLabel = 'Todos' } = {}) {
  if (!selectEl) {
    return;
  }
  selectEl.innerHTML = '';
  if (includeAll) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = allLabel;
    selectEl.appendChild(opt);
  }
  (options || []).forEach((item) => {
    const opt = document.createElement('option');
    opt.value = String(item.value ?? '');
    opt.textContent = String(item.label ?? '');
    selectEl.appendChild(opt);
  });
}

function getSetorLabelById(id) {
  const safeId = String(id || '').trim();
  const setor = (state.rhSetores || []).find((s) => String(s.id) === safeId);
  return setor ? String(setor.nome_setor || setor.nome || '').trim() : '';
}

function canUseRhModule() {
  const ctx = state.rhContext || null;
  return !!(ctx && ctx.ok);
}

async function loadRhContext() {
  const username = getAuthUsername();
  if (!username) {
    throw new Error('Usuário logado não identificado.');
  }
  const ctx = await window.api.rhGetContext({ username });
  state.rhContext = ctx || null;
  if (!ctx || !ctx.ok) {
    const err = (ctx && ctx.error) || 'Seu usuário não está configurado para RH.';
    throw new Error(err);
  }
  return ctx;
}

async function loadRhBaseLists() {
  const username = getAuthUsername();
  state.rhSetores = (await window.api.rhListSetores({ username }))?.setores || [];
  state.rhColaboradores = (await window.api.rhListColaboradores({ username }))?.colaboradores || [];
}

function refreshRhFilterOptions(selected = {}) {
  const setores = (state.rhSetores || []).map((s) => ({
    value: s.id,
    label: s.nome_setor || s.nome || `Setor ${s.id}`
  }));
  setRhSelectOptions(elements.rhFilterSector, setores, { includeAll: true, allLabel: 'Todos os setores' });
  if (elements.rhFilterSector && selected && selected.setorId != null) {
    elements.rhFilterSector.value = String(selected.setorId || '').trim();
  }

  const collabs = (state.rhColaboradores || []).map((c) => ({
    value: c.id,
    label: `${c.nome || 'Colaborador'}${c.setor_id ? ` • ${getSetorLabelById(c.setor_id)}` : ''}`
  }));
  setRhSelectOptions(elements.rhFilterCollab, collabs, { includeAll: true, allLabel: 'Todos os colaboradores' });
  if (elements.rhFilterCollab && selected && selected.colaboradorId != null) {
    elements.rhFilterCollab.value = String(selected.colaboradorId || '').trim();
  }
}

function openRhModal(modalEl) {
  if (modalEl) {
    modalEl.hidden = false;
  }
}

function closeRhModal(modalEl) {
  if (modalEl) {
    modalEl.hidden = true;
  }
}

function confirmRhAction(message, onConfirm) {
  if (!elements.rhConfirmModal) {
    return;
  }
  state.rhPendingConfirm = typeof onConfirm === 'function' ? onConfirm : null;
  if (elements.rhConfirmError) {
    elements.rhConfirmError.textContent = '';
  }
  if (elements.rhConfirmMessage) {
    elements.rhConfirmMessage.textContent = String(message || 'Confirme para continuar.');
  }
  openRhModal(elements.rhConfirmModal);
}

function formatHours(value) {
  const hours = Number(value || 0);
  const totalMinutes = Math.round(Math.abs(hours) * 60);
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  const sign = hours < 0 ? '-' : '';
  return `${sign}${h}:${String(m).padStart(2, '0')}`;
}

function parseDurationTextToHours(value) {
  const raw = String(value ?? '').trim();
  if (!raw) {
    return null;
  }

  const normalized = raw.replace(',', '.').replace(/\s+/g, '');
  const matchHm = normalized.match(/^(\d{1,3}):(\d{1,2})$/);
  if (matchHm) {
    const h = Number(matchHm[1]);
    const m = Number(matchHm[2]);
    if (!Number.isFinite(h) || !Number.isFinite(m) || h < 0 || m < 0 || m > 59) {
      return null;
    }
    return h + (m / 60);
  }

  const matchDec = normalized.match(/^\d{1,3}(\.\d+)?$/);
  if (matchDec) {
    const dec = Number(normalized);
    if (!Number.isFinite(dec) || dec < 0) {
      return null;
    }
    return dec;
  }

  return null;
}

function renderRhDashboard() {
  if (!elements.rhViewDashboard) {
    return;
  }
  elements.rhViewDashboard.innerHTML = '';

  if (!canUseRhModule()) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Seu usuário não está configurado no RH.';
    elements.rhViewDashboard.appendChild(p);
    return;
  }

  const kpis = state.rhKpis || null;
  if (!kpis) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Clique em Atualizar para carregar indicadores.';
    elements.rhViewDashboard.appendChild(p);
    return;
  }

  const grid = document.createElement('div');
  grid.className = 'rh-grid-2';

  const cards = [
    {
      title: 'Total de H.E no mês',
      value: formatHours(kpis.totalHoras || 0),
      sub: (kpis.comparacao && kpis.comparacao.prevMonth)
        ? `Vs ${kpis.comparacao.prevMonth}: ${kpis.comparacao.diff >= 0 ? '+' : ''}${formatHours(kpis.comparacao.diff || 0)} (${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Number(kpis.comparacao.pct || 0))}%)`
        : ''
    },
    {
      title: 'Média diária (dias úteis)',
      value: formatHours(kpis.mediaDiaria || 0),
      sub: `${formatInteger(kpis.diasTrabalhadosNoMes || 0)} dia(s) útil(eis) | ${formatInteger(kpis.diasComLancamento || 0)} com lançamento`
    },
    { title: 'Colaboradores no período', value: formatInteger(kpis.colaboradoresNoPeriodo || 0), sub: `${formatInteger(kpis.totalRegistros || 0)} lançamento(s)` },
    {
      title: 'Alertas de limite',
      value: formatInteger((kpis.alertas || []).length),
      sub: kpis.limiteTexto || ''
    },
    {
      title: 'Margem média (ponderada)',
      value: `${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Number(kpis.margemMediaPonderada || 0))}%`,
      sub: 'Faixa por total mensal'
    }
  ];

  cards.forEach((card) => {
    const el = document.createElement('div');
    el.className = 'rh-kpi-card';
    const title = document.createElement('p');
    title.className = 'rh-kpi-title';
    title.textContent = card.title;
    const value = document.createElement('p');
    value.className = 'rh-kpi-value';
    value.textContent = card.value;
    const sub = document.createElement('p');
    sub.className = 'rh-kpi-sub';
    sub.textContent = card.sub;
    el.appendChild(title);
    el.appendChild(value);
    el.appendChild(sub);
    grid.appendChild(el);
  });
  elements.rhViewDashboard.appendChild(grid);

  const sections = document.createElement('div');
  sections.className = 'rh-grid-2';

  const buildBarSection = (titleText, rows, labelKey) => {
    const card = document.createElement('div');
    card.className = 'rh-kpi-card';
    const t = document.createElement('p');
    t.className = 'rh-kpi-title';
    t.textContent = titleText;
    card.appendChild(t);
    const list = document.createElement('div');
    list.className = 'rh-barlist';
    const safeRows = Array.isArray(rows) ? rows : [];
    const max = Math.max(1, ...safeRows.map((r) => Number(r.horas || 0)));
    safeRows.slice(0, 10).forEach((row) => {
      const line = document.createElement('div');
      line.className = 'rh-barrow';
      const label = document.createElement('div');
      label.className = 'rh-barlabel';
      const baseLabel = String(row[labelKey] || '—');
      if (labelKey === 'colaborador' && row.margem_percent != null) {
        const pct = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 }).format(Number(row.margem_percent || 0));
        label.textContent = `${baseLabel} (${pct}%)`;
      } else {
        label.textContent = baseLabel;
      }
      const track = document.createElement('div');
      track.className = 'rh-bartrack';
      const fill = document.createElement('div');
      fill.className = 'rh-barfill';
      fill.style.width = `${Math.max(0, Math.min(100, (Number(row.horas || 0) / max) * 100))}%`;
      track.appendChild(fill);
      const value = document.createElement('div');
      value.className = 'rh-barvalue';
      value.textContent = formatHours(row.horas || 0);
      line.appendChild(label);
      line.appendChild(track);
      line.appendChild(value);
      list.appendChild(line);
    });
    if (!safeRows.length) {
      const empty = document.createElement('p');
      empty.className = 'pcp-empty-state';
      empty.textContent = 'Sem lançamentos no período.';
      list.appendChild(empty);
    }
    card.appendChild(list);
    return card;
  };

  sections.appendChild(buildBarSection('H.E por setor', kpis.porSetor || [], 'setor'));
  sections.appendChild(buildBarSection('Ranking de colaboradores', kpis.rankingColaboradores || [], 'colaborador'));
  elements.rhViewDashboard.appendChild(sections);

  const evo = Array.isArray(kpis.evolucaoMensal) ? kpis.evolucaoMensal : [];
  if (evo.length) {
    const evoCard = buildBarSection('Evolução mensal (últimos 6 meses)', evo.map((r) => ({
      month: r.month,
      horas: r.totalHoras
    })), 'month');
    elements.rhViewDashboard.appendChild(evoCard);
  }
}

function renderRhRegistros() {
  if (!elements.rhViewRegistros) {
    return;
  }
  elements.rhViewRegistros.innerHTML = '';

  const tableWrap = document.createElement('div');
  tableWrap.className = 'table-wrapper pcp-table-wrapper';
  const table = document.createElement('table');
  table.className = 'pcp-table';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  thead.innerHTML = '<tr><th>Data</th><th>Colaborador</th><th>Setor</th><th>Horas</th><th>Tipo</th><th>Observação</th><th>Criado por</th><th>Ações</th></tr>';

  const collabById = new Map((state.rhColaboradores || []).map((c) => [String(c.id), c]));
  const rows = state.rhHorasExtras || [];
  rows.forEach((r) => {
    const tr = document.createElement('tr');
    const collab = collabById.get(String(r.colaborador_id)) || null;
    const tdData = document.createElement('td');
    tdData.textContent = formatYmdToBr(String(r.data || '').trim());
    const tdColab = document.createElement('td');
    tdColab.textContent = collab ? String(collab.nome || '').trim() : '—';
    const tdSetor = document.createElement('td');
    tdSetor.textContent = collab ? (getSetorLabelById(collab.setor_id) || '—') : '—';
    const tdHoras = document.createElement('td');
    tdHoras.textContent = formatHours(r.quantidade_horas || 0);
    const tdTipo = document.createElement('td');
    tdTipo.textContent = String(r.tipo_hora || '').trim();
    const tdObs = document.createElement('td');
    tdObs.textContent = String(r.observacao || '').trim();
    const tdUser = document.createElement('td');
    tdUser.textContent = String(r.criado_por || '').trim();
    const tdAcoes = document.createElement('td');
    if (canEditRh()) {
      const edit = document.createElement('button');
      edit.className = 'btn secondary';
      edit.type = 'button';
      edit.textContent = 'Editar';
      edit.addEventListener('click', () => openRhHeModal(r));
      const del = document.createElement('button');
      del.className = 'btn danger';
      del.type = 'button';
      del.style.marginLeft = '8px';
      del.textContent = 'Apagar';
      del.addEventListener('click', () => confirmRhAction('Apagar este lançamento de H.E?', async () => {
        await window.api.rhDeleteHoraExtra({ username: getAuthUsername(), id: r.id });
        await refreshRhAll();
        setRhHeSubView('registros');
      }));
      tdAcoes.appendChild(edit);
      tdAcoes.appendChild(del);
    } else {
      tdAcoes.textContent = '—';
    }

    tr.appendChild(tdData);
    tr.appendChild(tdColab);
    tr.appendChild(tdSetor);
    tr.appendChild(tdHoras);
    tr.appendChild(tdTipo);
    tr.appendChild(tdObs);
    tr.appendChild(tdUser);
    tr.appendChild(tdAcoes);
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  elements.rhViewRegistros.appendChild(tableWrap);

  if (!rows.length) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Sem lançamentos no período.';
    elements.rhViewRegistros.appendChild(p);
  }
}

function renderRhAtrasosDashboard() {
  if (!elements.rhViewAtrasos) {
    return;
  }
  elements.rhViewAtrasos.innerHTML = '';

  if (!canUseRhModule()) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Seu usuário não está configurado no RH.';
    elements.rhViewAtrasos.appendChild(p);
    return;
  }

  const kpis = state.rhAtrasosKpis || null;
  if (!kpis) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Clique em Atualizar para carregar indicadores.';
    elements.rhViewAtrasos.appendChild(p);
    return;
  }

  const grid = document.createElement('div');
  grid.className = 'rh-grid-2';

  const cards = [
    {
      title: 'Total de atrasos no mês',
      value: formatHours(kpis.totalHoras || 0),
      sub: (kpis.comparacao && kpis.comparacao.prevMonth)
        ? `Vs ${kpis.comparacao.prevMonth}: ${kpis.comparacao.diff >= 0 ? '+' : ''}${formatHours(kpis.comparacao.diff || 0)} (${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Number(kpis.comparacao.pct || 0))}%)`
        : ''
    },
    {
      title: 'Média diária (dias úteis)',
      value: formatHours(kpis.mediaDiaria || 0),
      sub: `${formatInteger(kpis.diasTrabalhadosNoMes || 0)} dia(s) útil(eis) | ${formatInteger(kpis.diasComLancamento || 0)} com atraso`
    },
    {
      title: 'Colaboradores com atraso',
      value: formatInteger(kpis.colaboradoresNoPeriodo || 0),
      sub: `${formatInteger(kpis.totalRegistros || 0)} registro(s)`
    }
  ];

  cards.forEach((card) => {
    const el = document.createElement('div');
    el.className = 'rh-kpi-card';
    const title = document.createElement('p');
    title.className = 'rh-kpi-title';
    title.textContent = card.title;
    const value = document.createElement('p');
    value.className = 'rh-kpi-value';
    value.textContent = card.value;
    const sub = document.createElement('p');
    sub.className = 'rh-kpi-sub';
    sub.textContent = card.sub;
    el.appendChild(title);
    el.appendChild(value);
    el.appendChild(sub);
    grid.appendChild(el);
  });
  elements.rhViewAtrasos.appendChild(grid);

  const sections = document.createElement('div');
  sections.className = 'rh-grid-2';

  const buildBarSection = (titleText, rows, labelKey) => {
    const card = document.createElement('div');
    card.className = 'rh-kpi-card';
    const t = document.createElement('p');
    t.className = 'rh-kpi-title';
    t.textContent = titleText;
    card.appendChild(t);
    const list = document.createElement('div');
    list.className = 'rh-barlist';
    const safeRows = Array.isArray(rows) ? rows : [];
    const max = Math.max(1, ...safeRows.map((r) => Number(r.horas || 0)));
    safeRows.slice(0, 10).forEach((row) => {
      const line = document.createElement('div');
      line.className = 'rh-barrow';
      const label = document.createElement('div');
      label.className = 'rh-barlabel';
      label.textContent = String(row[labelKey] || '—');
      const track = document.createElement('div');
      track.className = 'rh-bartrack';
      const fill = document.createElement('div');
      fill.className = 'rh-barfill';
      fill.style.width = `${Math.max(0, Math.min(100, (Number(row.horas || 0) / max) * 100))}%`;
      track.appendChild(fill);
      const value = document.createElement('div');
      value.className = 'rh-barvalue';
      value.textContent = formatHours(row.horas || 0);
      line.appendChild(label);
      line.appendChild(track);
      line.appendChild(value);
      list.appendChild(line);
    });
    if (!safeRows.length) {
      const empty = document.createElement('p');
      empty.className = 'pcp-empty-state';
      empty.textContent = 'Sem atrasos no período.';
      list.appendChild(empty);
    }
    card.appendChild(list);
    return card;
  };

  sections.appendChild(buildBarSection('Atrasos por setor', kpis.porSetor || [], 'setor'));
  sections.appendChild(buildBarSection('Ranking de colaboradores', kpis.rankingColaboradores || [], 'colaborador'));
  elements.rhViewAtrasos.appendChild(sections);

  const evo = Array.isArray(kpis.evolucaoMensal) ? kpis.evolucaoMensal : [];
  if (evo.length) {
    const evoCard = buildBarSection('Evolução mensal (últimos 6 meses)', evo.map((r) => ({
      month: r.month,
      horas: r.totalHoras
    })), 'month');
    elements.rhViewAtrasos.appendChild(evoCard);
  }
}

async function refreshRhAll() {
  const prevFilters = getRhUiFilters();
  await loadRhContext();
  await loadRhBaseLists();
  refreshRhFilterOptions(prevFilters);
  if (elements.rhFilterMonth && prevFilters.month) {
    elements.rhFilterMonth.value = prevFilters.month;
  }
  if (elements.rhFilterType) {
    elements.rhFilterType.value = String(prevFilters.tipoHora || '').trim();
  }
  const username = getAuthUsername();
  const filters = getRhUiFilters();
  state.rhKpis = (await window.api.rhGetKpis({ username, ...filters }))?.kpis || null;
  state.rhHorasExtras = (await window.api.rhListHorasExtras({ username, ...filters }))?.rows || [];
  state.rhAtrasosKpis = (await window.api.rhGetAtrasosKpis({ username, month: filters.month, setorId: filters.setorId, colaboradorId: filters.colaboradorId }))?.kpis || null;
  state.rhAtrasosRows = (await window.api.rhListAtrasos({ username, month: filters.month, setorId: filters.setorId, colaboradorId: filters.colaboradorId }))?.rows || [];
  renderRhDashboard();
  renderRhAtrasosDashboard();
  renderRhRegistros();
  renderRhColaboradores();
  renderRhSetores();
  if (state.rhHeSubView === 'config') {
    await renderRhConfig();
  }
  if (state.rhHeSubView === 'audit') {
    await renderRhAudit();
  }
}

function openRhSectorModal(seed) {
  if (!elements.rhSectorModal) {
    return;
  }
  if (elements.rhSectorError) {
    elements.rhSectorError.textContent = '';
  }
  if (elements.rhSectorId) {
    elements.rhSectorId.value = seed && seed.id ? String(seed.id) : '';
  }
  if (elements.rhSectorNome) {
    elements.rhSectorNome.value = seed ? String(seed.nome_setor || '') : '';
  }
  if (elements.rhSectorGestor) {
    elements.rhSectorGestor.value = seed ? String(seed.gestor_responsavel || '') : '';
  }
  if (elements.rhSectorCc) {
    elements.rhSectorCc.value = seed ? String(seed.centro_custo || '') : '';
  }
  openRhModal(elements.rhSectorModal);
  elements.rhSectorNome?.focus?.();
}

function closeRhSectorModal() {
  closeRhModal(elements.rhSectorModal);
}

function openRhCollabModal(seed) {
  if (!elements.rhCollabModal) {
    return;
  }
  if (elements.rhCollabError) {
    elements.rhCollabError.textContent = '';
  }
  if (elements.rhCollabId) {
    elements.rhCollabId.value = seed && seed.id ? String(seed.id) : '';
  }
  if (elements.rhCollabNome) {
    elements.rhCollabNome.value = seed ? String(seed.nome || '') : '';
  }
  if (elements.rhCollabMatricula) {
    elements.rhCollabMatricula.value = seed ? String(seed.matricula || '') : '';
  }
  if (elements.rhCollabSetor) {
    setRhSelectOptions(
      elements.rhCollabSetor,
      (state.rhSetores || []).map((s) => ({ value: s.id, label: s.nome_setor })),
      { includeAll: true, allLabel: 'Selecione' }
    );
    elements.rhCollabSetor.value = seed && seed.setor_id ? String(seed.setor_id) : '';
  }
  if (elements.rhCollabCargo) {
    elements.rhCollabCargo.value = seed ? String(seed.cargo || '') : '';
  }
  if (elements.rhCollabAdmissao) {
    elements.rhCollabAdmissao.value = seed ? String(seed.data_admissao || '') : '';
  }
  if (elements.rhCollabStatus) {
    elements.rhCollabStatus.value = seed ? String(seed.status || 'ativo') : 'ativo';
  }
  if (elements.rhCollabLimite) {
    const raw = seed ? String(seed.limite_mensal || '').trim() : '';
    elements.rhCollabLimite.value = raw;
  }
  openRhModal(elements.rhCollabModal);
  elements.rhCollabNome?.focus?.();
}

function closeRhCollabModal() {
  closeRhModal(elements.rhCollabModal);
}

function openRhHeModal(seed) {
  if (!elements.rhHeModal) {
    return;
  }
  if (elements.rhHeError) {
    elements.rhHeError.textContent = '';
  }
  if (elements.rhHeId) {
    elements.rhHeId.value = seed && seed.id ? String(seed.id) : '';
  }
  setRhHeModalKind('he', { lock: !!(seed && seed.id) });
  if (elements.rhHeColaborador) {
    const active = (state.rhColaboradores || []).filter((c) => String(c.status || '').trim() !== 'inativo');
    setRhSelectOptions(
      elements.rhHeColaborador,
      active.map((c) => ({ value: c.id, label: c.nome })),
      { includeAll: false }
    );
    elements.rhHeColaborador.value = seed && seed.colaborador_id
      ? String(seed.colaborador_id)
      : String(active[0]?.id || '');
  }
  if (elements.rhHeData) {
    elements.rhHeData.value = seed && seed.data ? String(seed.data) : getTodayYmdLocal();
  }
  if (elements.rhHeQtd) {
    elements.rhHeQtd.value = seed && seed.quantidade_horas !== undefined ? formatHours(seed.quantidade_horas) : '';
  }
  if (elements.rhHeTipo) {
    elements.rhHeTipo.value = seed ? String(seed.tipo_hora || '50%') : '50%';
  }
  if (elements.rhHeObs) {
    elements.rhHeObs.value = seed ? String(seed.observacao || '') : '';
  }
  if (elements.rhHeJust) {
    elements.rhHeJust.value = seed ? String(seed.justificativa || '') : '';
  }
  openRhModal(elements.rhHeModal);
  elements.rhHeQtd?.focus?.();
}

function closeRhHeModal() {
  closeRhModal(elements.rhHeModal);
}

function setRhHeModalKind(nextKind, { lock = false } = {}) {
  const kind = String(nextKind || '').trim() === 'atraso' ? 'atraso' : 'he';

  if (elements.rhHeKind) {
    elements.rhHeKind.value = kind;
    elements.rhHeKind.disabled = !!lock;
  }

  const isAtraso = kind === 'atraso';
  if (elements.rhHeTipoField) {
    elements.rhHeTipoField.hidden = isAtraso;
  }
  if (elements.rhHeJustField) {
    elements.rhHeJustField.hidden = isAtraso;
  }
  if (elements.rhHeQtdLabel) {
    elements.rhHeQtdLabel.textContent = isAtraso ? 'Tempo de atraso (hh:mm)' : 'Quantidade de horas (hh:mm)';
  }
  if (elements.rhHeTitle) {
    elements.rhHeTitle.textContent = isAtraso ? 'Lançamento de Atrasos' : 'Lançamento de Horas Extras';
  }
  if (elements.rhHeModalSubtitle) {
    elements.rhHeModalSubtitle.textContent = isAtraso
      ? 'Registre atrasos diários por colaborador.'
      : 'Registre horas extras diárias por colaborador.';
  }

  if (elements.rhHeQtd) {
    elements.rhHeQtd.placeholder = 'Ex: 1:30';
  }
}

function setRhImportWorkbookPath(nextPath) {
  const safePath = String(nextPath || '').trim();
  state.rhImportWorkbookPath = safePath;

  if (elements.rhImportFilePath) {
    const fileName = safePath ? safePath.split(/[/\\\\]/).filter(Boolean).slice(-1)[0] : '';
    elements.rhImportFilePath.value = fileName;
    elements.rhImportFilePath.title = safePath;
  }

  if (elements.rhImportConfirm) {
    elements.rhImportConfirm.disabled = !safePath;
  }
}

function openRhImportModal() {
  if (!elements.rhImportModal) {
    return;
  }
  if (elements.rhImportError) {
    elements.rhImportError.textContent = '';
  }
  if (elements.rhImportSource) {
    elements.rhImportSource.value = 'ponto_interno';
  }
  if (elements.rhImportDate) {
    elements.rhImportDate.value = getTodayYmdLocal();
  }
  if (elements.rhImportFile) {
    elements.rhImportFile.value = '';
  }
  setRhImportWorkbookPath('');
  openRhModal(elements.rhImportModal);
}

function closeRhImportModal() {
  closeRhModal(elements.rhImportModal);
}

async function handleRhImportConfirm() {
  try {
    const username = getAuthUsername();
    if (!username) {
      throw new Error('Usuário não autenticado.');
    }
    if (!canEditRh()) {
      throw new Error('Somente RH (edição) pode importar.');
    }

    const source = String(elements.rhImportSource?.value || '').trim();
    if (source !== 'ponto_interno') {
      throw new Error('Importação de Ponto cartão ainda não disponível.');
    }

    const date = String(elements.rhImportDate?.value || '').trim();
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      throw new Error('Informe a data do lançamento.');
    }

    const fallbackFilePath = elements.rhImportFile?.files?.[0]?.path ? String(elements.rhImportFile.files[0].path) : '';
    const filePath = String(state.rhImportWorkbookPath || '').trim() || fallbackFilePath;
    if (!filePath) {
      throw new Error('Selecione a planilha base.');
    }

    if (elements.rhImportConfirm) {
      elements.rhImportConfirm.disabled = true;
    }
    if (elements.rhImportError) {
      elements.rhImportError.textContent = '';
    }

    const result = await window.api.rhImportInterno({ username, date, workbookPath: filePath });
    const imported = Number(result?.imported || 0);
    const updated = Number(result?.updated || 0);
    const unmatched = Number(result?.unmatched || 0);
    const skipped = Number(result?.skipped || 0);
    const importedAtrasos = Number(result?.importedAtrasos || 0);
    const updatedAtrasos = Number(result?.updatedAtrasos || 0);
    const unmatchedAtrasos = Number(result?.unmatchedAtrasos || 0);
    const skippedAtrasos = Number(result?.skippedAtrasos || 0);

    showToast(`H.E — Importado: ${imported} | Atualizado: ${updated} | Sem match: ${unmatched} | Zerado/vazio: ${skipped} • Atrasos — Importado: ${importedAtrasos} | Atualizado: ${updatedAtrasos} | Sem match: ${unmatchedAtrasos} | Zerado/vazio: ${skippedAtrasos}`);

    const unmatchedSheetsHe = result?.details?.unmatchedSheets || [];
    const unmatchedSheetsAtrasos = result?.details?.unmatchedSheetsAtrasos || [];
    if ((unmatched || unmatchedAtrasos) && elements.rhImportError) {
      const namesHe = unmatchedSheetsHe.slice(0, 8).map((x) => x.sheetName).filter(Boolean);
      const namesAtr = unmatchedSheetsAtrasos.slice(0, 8).map((x) => x.sheetName).filter(Boolean);
      const msgParts = [];
      if (unmatched && namesHe.length) {
        msgParts.push(`H.E sem match: ${namesHe.join(' | ')}${unmatchedSheetsHe.length > 8 ? ` | +${unmatchedSheetsHe.length - 8}` : ''}`);
      } else if (unmatched) {
        msgParts.push('H.E: existem abas sem match (verifique nomes).');
      }
      if (unmatchedAtrasos && namesAtr.length) {
        msgParts.push(`Atrasos sem match: ${namesAtr.join(' | ')}${unmatchedSheetsAtrasos.length > 8 ? ` | +${unmatchedSheetsAtrasos.length - 8}` : ''}`);
      } else if (unmatchedAtrasos) {
        msgParts.push('Atrasos: existem abas sem match (verifique nomes).');
      }
      elements.rhImportError.textContent = msgParts.join(' • ');
      return;
    }

    closeRhImportModal();
    await refreshRhAll();
  } catch (error) {
    if (elements.rhImportError) {
      elements.rhImportError.textContent = error.message || 'Erro ao importar.';
    } else {
      showToast(error.message || 'Erro ao importar.', true);
    }
  } finally {
    if (elements.rhImportConfirm) {
      elements.rhImportConfirm.disabled = false;
    }
  }
}

function renderRhSetores() {
  if (!elements.rhViewSetores) {
    return;
  }
  elements.rhViewSetores.innerHTML = '';
  const header = document.createElement('div');
  header.className = 'pcp-top-actions';
  if (canEditRh()) {
    const btn = document.createElement('button');
    btn.className = 'btn secondary';
    btn.type = 'button';
    btn.textContent = 'Novo setor';
    btn.addEventListener('click', () => openRhSectorModal(null));
    header.appendChild(btn);
  }
  elements.rhViewSetores.appendChild(header);

  const tableWrap = document.createElement('div');
  tableWrap.className = 'table-wrapper pcp-table-wrapper';
  const table = document.createElement('table');
  table.className = 'pcp-table';
  table.innerHTML = '<thead><tr><th>Setor</th><th>Gestor</th><th>Centro custo</th><th>Ações</th></tr></thead>';
  const tbody = document.createElement('tbody');

  (state.rhSetores || []).forEach((s) => {
    const tr = document.createElement('tr');
    const tdNome = document.createElement('td');
    tdNome.textContent = String(s.nome_setor || '').trim();
    const tdGestor = document.createElement('td');
    tdGestor.textContent = String(s.gestor_responsavel || '').trim();
    const tdCc = document.createElement('td');
    tdCc.textContent = String(s.centro_custo || '').trim();
    const tdAcoes = document.createElement('td');
    if (canEditRh()) {
      const edit = document.createElement('button');
      edit.className = 'btn secondary';
      edit.type = 'button';
      edit.textContent = 'Editar';
      edit.addEventListener('click', () => openRhSectorModal(s));
      const del = document.createElement('button');
      del.className = 'btn danger';
      del.type = 'button';
      del.style.marginLeft = '8px';
      del.textContent = 'Inativar';
      del.addEventListener('click', () => confirmRhAction(`Inativar o setor "${s.nome_setor}"?`, async () => {
        await window.api.rhDeleteSetor({ username: getAuthUsername(), id: s.id });
        await refreshRhAll();
        setRhHeSubView('setores');
      }));
      tdAcoes.appendChild(edit);
      tdAcoes.appendChild(del);
    } else {
      tdAcoes.textContent = '—';
    }
    tr.appendChild(tdNome);
    tr.appendChild(tdGestor);
    tr.appendChild(tdCc);
    tr.appendChild(tdAcoes);
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  elements.rhViewSetores.appendChild(tableWrap);
}

function renderRhColaboradores() {
  if (!elements.rhViewColaboradores) {
    return;
  }
  elements.rhViewColaboradores.innerHTML = '';
  const header = document.createElement('div');
  header.className = 'pcp-top-actions';
  if (canEditRh()) {
    const btn = document.createElement('button');
    btn.className = 'btn secondary';
    btn.type = 'button';
    btn.textContent = 'Novo colaborador';
    btn.addEventListener('click', () => openRhCollabModal(null));
    header.appendChild(btn);
  }
  elements.rhViewColaboradores.appendChild(header);

  const tableWrap = document.createElement('div');
  tableWrap.className = 'table-wrapper pcp-table-wrapper';
  const table = document.createElement('table');
  table.className = 'pcp-table';
  table.innerHTML = '<thead><tr><th>Nome</th><th>Setor</th><th>Cargo</th><th>Status</th><th>Admissão</th><th>Limite</th><th>Ações</th></tr></thead>';
  const tbody = document.createElement('tbody');

  (state.rhColaboradores || []).forEach((c) => {
    const tr = document.createElement('tr');
    const tdNome = document.createElement('td');
    tdNome.textContent = String(c.nome || '').trim();
    const tdSetor = document.createElement('td');
    tdSetor.textContent = getSetorLabelById(c.setor_id) || '—';
    const tdCargo = document.createElement('td');
    tdCargo.textContent = String(c.cargo || '').trim();
    const tdStatus = document.createElement('td');
    tdStatus.textContent = String(c.status || '').trim();
    const tdAdm = document.createElement('td');
    tdAdm.textContent = formatYmdToBr(String(c.data_admissao || '').trim());
    const tdLim = document.createElement('td');
    tdLim.textContent = String(c.limite_mensal || '').trim() ? `${String(c.limite_mensal).trim()} h` : '—';
    const tdAcoes = document.createElement('td');
    if (canEditRh()) {
      const edit = document.createElement('button');
      edit.className = 'btn secondary';
      edit.type = 'button';
      edit.textContent = 'Editar';
      edit.addEventListener('click', () => openRhCollabModal(c));
      const toggle = document.createElement('button');
      toggle.className = 'btn secondary';
      toggle.type = 'button';
      toggle.style.marginLeft = '8px';
      toggle.textContent = String(c.status || '').trim() === 'inativo' ? 'Ativar' : 'Inativar';
      toggle.addEventListener('click', () => confirmRhAction(`Alterar status de "${c.nome}"?`, async () => {
        await window.api.rhToggleColaboradorStatus({ username: getAuthUsername(), id: c.id });
        await refreshRhAll();
        setRhHeSubView('colaboradores');
      }));
      const del = document.createElement('button');
      del.className = 'btn danger';
      del.type = 'button';
      del.style.marginLeft = '8px';
      del.textContent = 'Remover';
      del.addEventListener('click', () => confirmRhAction(`Remover o colaborador "${c.nome}"?`, async () => {
        await window.api.rhDeleteColaborador({ username: getAuthUsername(), id: c.id });
        await refreshRhAll();
        setRhHeSubView('colaboradores');
      }));
      tdAcoes.appendChild(edit);
      tdAcoes.appendChild(toggle);
      tdAcoes.appendChild(del);
    } else {
      tdAcoes.textContent = '—';
    }
    tr.appendChild(tdNome);
    tr.appendChild(tdSetor);
    tr.appendChild(tdCargo);
    tr.appendChild(tdStatus);
    tr.appendChild(tdAdm);
    tr.appendChild(tdLim);
    tr.appendChild(tdAcoes);
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  elements.rhViewColaboradores.appendChild(tableWrap);
}

async function renderRhConfig() {
  if (!elements.rhViewConfig) {
    return;
  }
  elements.rhViewConfig.innerHTML = '';

  const username = getAuthUsername();
  const canEdit = canEditRh();
  const cfg = (await window.api.rhConfigGet({ username }))?.config || {};
  const ctx = state.rhContext || {};

  const card = document.createElement('div');
  card.className = 'rh-kpi-card';
  const title = document.createElement('p');
  title.className = 'rh-kpi-title';
  title.textContent = 'Configurações';
  card.appendChild(title);

  const grid = document.createElement('div');
  grid.className = 'rh-grid-2';

  const buildNumberField = (labelText, defaultValue) => {
    const field = document.createElement('div');
    field.className = 'field';
    const label = document.createElement('label');
    label.textContent = labelText;
    const input = document.createElement('input');
    input.type = 'number';
    input.step = '0.25';
    input.min = '0';
    input.value = String(defaultValue);
    input.disabled = !canEdit;
    field.appendChild(label);
    field.appendChild(input);
    return { field, input };
  };

  const justField = buildNumberField('Justificativa obrigatória acima de (horas)', cfg.justificativa_acima_horas ?? 2);
  const limField = buildNumberField('Limite mensal padrão por colaborador (horas)', cfg.limite_mensal_padrao ?? 40);
  grid.appendChild(justField.field);
  grid.appendChild(limField.field);

  const m1h = buildNumberField('Margem faixa 1 até (horas)', cfg.margem_faixa_1_ate_horas ?? 25);
  const m1p = buildNumberField('Margem faixa 1 (%)', cfg.margem_faixa_1_percent ?? 50);
  const m2h = buildNumberField('Margem faixa 2 até (horas)', cfg.margem_faixa_2_ate_horas ?? 40);
  const m2p = buildNumberField('Margem faixa 2 (%)', cfg.margem_faixa_2_percent ?? 60);
  const m3h = buildNumberField('Margem faixa 3 até (horas)', cfg.margem_faixa_3_ate_horas ?? 60);
  const m3p = buildNumberField('Margem faixa 3 (%)', cfg.margem_faixa_3_percent ?? 80);
  const m4p = buildNumberField('Margem faixa 4 (%) (acima)', cfg.margem_faixa_4_percent ?? 100);
  grid.appendChild(m1h.field);
  grid.appendChild(m1p.field);
  grid.appendChild(m2h.field);
  grid.appendChild(m2p.field);
  grid.appendChild(m3h.field);
  grid.appendChild(m3p.field);
  grid.appendChild(m4p.field);
  card.appendChild(grid);

  if (canEdit) {
    const actions = document.createElement('div');
    actions.className = 'pcp-top-actions';
    const btn = document.createElement('button');
    btn.className = 'btn primary';
    btn.type = 'button';
    btn.textContent = 'Salvar configurações';
    btn.addEventListener('click', async () => {
      try {
        await window.api.rhConfigSet({
          username,
          config: {
            justificativa_acima_horas: Number(justField.input.value || 0),
            limite_mensal_padrao: Number(limField.input.value || 0),
            margem_faixa_1_ate_horas: Number(m1h.input.value || 0),
            margem_faixa_1_percent: Number(m1p.input.value || 0),
            margem_faixa_2_ate_horas: Number(m2h.input.value || 0),
            margem_faixa_2_percent: Number(m2p.input.value || 0),
            margem_faixa_3_ate_horas: Number(m3h.input.value || 0),
            margem_faixa_3_percent: Number(m3p.input.value || 0),
            margem_faixa_4_percent: Number(m4p.input.value || 0)
          }
        });
        showToast('Configurações RH salvas.');
        await refreshRhAll();
        setRhHeSubView('config');
      } catch (error) {
        showToast(error.message || 'Erro ao salvar configurações RH.', true);
      }
    });
    actions.appendChild(btn);
    card.appendChild(actions);
  } else {
    const note = document.createElement('p');
    note.className = 'rh-kpi-sub';
    note.textContent = 'Somente RH pode alterar configurações.';
    card.appendChild(note);
  }

  elements.rhViewConfig.appendChild(card);

  const infoCard = document.createElement('div');
  infoCard.className = 'rh-kpi-card';
  const infoTitle = document.createElement('p');
  infoTitle.className = 'rh-kpi-title';
  infoTitle.textContent = 'Permissões e Base';
  const infoText = document.createElement('p');
  infoText.className = 'rh-kpi-sub';
  infoText.textContent = 'O acesso ao RH é controlado na aba Configuração (Logins): "RH" para visualizar e "RH Editar" para lançar/editar/apagar.';
  const baseText = document.createElement('p');
  baseText.className = 'rh-kpi-sub';
  baseText.textContent = ctx.dbFilePath ? `Base: ${ctx.dbFilePath}` : '';
  infoCard.appendChild(infoTitle);
  infoCard.appendChild(infoText);
  if (baseText.textContent) {
    infoCard.appendChild(baseText);
  }
  elements.rhViewConfig.appendChild(infoCard);
}

async function renderRhAudit() {
  if (!elements.rhViewAudit) {
    return;
  }
  elements.rhViewAudit.innerHTML = '';
  const username = getAuthUsername();
  const audit = (await window.api.rhAuditList({ username, limit: 200 }))?.rows || [];
  state.rhAudit = audit;

  const tableWrap = document.createElement('div');
  tableWrap.className = 'table-wrapper pcp-table-wrapper';
  const table = document.createElement('table');
  table.className = 'pcp-table';
  table.innerHTML = '<thead><tr><th>Data/Hora</th><th>Usuário</th><th>Ação</th><th>Entidade</th><th>ID</th></tr></thead>';
  const tbody = document.createElement('tbody');
  audit.forEach((row) => {
    const tr = document.createElement('tr');
    const tdDt = document.createElement('td');
    tdDt.textContent = String(row.data_hora || '').replace('T', ' ').slice(0, 19);
    const tdU = document.createElement('td');
    tdU.textContent = String(row.username || '').trim();
    const tdA = document.createElement('td');
    tdA.textContent = String(row.action || '').trim();
    const tdE = document.createElement('td');
    tdE.textContent = String(row.entity || '').trim();
    const tdId = document.createElement('td');
    tdId.textContent = String(row.entity_id || '').trim();
    tr.appendChild(tdDt);
    tr.appendChild(tdU);
    tr.appendChild(tdA);
    tr.appendChild(tdE);
    tr.appendChild(tdId);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  elements.rhViewAudit.appendChild(tableWrap);

  if (!audit.length) {
    const p = document.createElement('p');
    p.className = 'pcp-empty-state';
    p.textContent = 'Sem auditoria no momento.';
    elements.rhViewAudit.appendChild(p);
  }
}

function getVersionParts(version) {
  const match = String(version).trim().match(/^(\d+)\.(\d+)(?:\.(\d+))?$/);
  if (!match) {
    return [0, 0, 0];
  }

  return [Number(match[1]), Number(match[2]), Number(match[3] || 0)];
}

function getSortedFeatureLogGroups() {
  return [...FEATURE_LOG_GROUPS].sort((a, b) => {
    const [aMajor, aMinor, aPatch] = getVersionParts(a.version);
    const [bMajor, bMinor, bPatch] = getVersionParts(b.version);

    if (bMajor !== aMajor) {
      return bMajor - aMajor;
    }
    if (bMinor !== aMinor) {
      return bMinor - aMinor;
    }
    return bPatch - aPatch;
  });
}

function updateFooterVersion() {
  if (!elements.footerVersion) {
    return;
  }
  const latest = getSortedFeatureLogGroups()[0] || null;
  if (!latest) {
    elements.footerVersion.textContent = 'Versão: - | Atualizado em: -';
    return;
  }
  const dateLabel = String(latest.date || '').trim() || '-';
  elements.footerVersion.textContent = `Versão: ${latest.version} | Atualizado em: ${dateLabel}`;
}

function renderLogUpdates() {
  if (!elements.logUpdatesContent) {
    return;
  }

  elements.logUpdatesContent.innerHTML = '';
  const sortedGroups = getSortedFeatureLogGroups();

  sortedGroups.forEach((group) => {
    const card = document.createElement('article');
    card.className = 'log-card';

    const title = document.createElement('h3');
    const versionLabel = group.date
      ? `Atualizacao ${group.version} - ${group.date}`
      : `Atualizacao ${group.version}`;
    title.textContent = group.title
      ? `${versionLabel} | ${group.title}`
      : versionLabel;
    card.appendChild(title);

    const list = document.createElement('ul');
    list.className = 'log-list';
    group.items.forEach((name) => {
      const item = document.createElement('li');
      item.textContent = name;
      list.appendChild(item);
    });

    card.appendChild(list);
    elements.logUpdatesContent.appendChild(card);
  });
}

function applyTheme(theme) {
  const safeTheme = theme === 'dark' ? 'dark' : 'light';
  elements.body.setAttribute('data-theme', safeTheme);
  if (elements.themeSwitch) {
    elements.themeSwitch.checked = safeTheme === 'dark';
  }
  if (elements.brandLogo) {
    const nextLogo = safeTheme === 'dark'
      ? elements.brandLogo.dataset.logoDark
      : elements.brandLogo.dataset.logoLight;
    if (nextLogo) {
      elements.brandLogo.src = nextLogo;
    }
  }
}

function formatInteger(value) {
  return new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function formatPercent(value) {
  return `${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Number(value || 0))}%`;
}

function formatWeightKg(value) {
  return `${new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 3, maximumFractionDigits: 3 }).format(Number(value || 0))} Kg`;
}

function getTodayYmdLocal() {
  const now = new Date();
  const offsetMs = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offsetMs).toISOString().slice(0, 10);
}

function formatYmdToBr(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return text;
  }
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function ensureDateValue() {
  if (elements.dateInput.value) {
    return elements.dateInput.value;
  }

  const today = getTodayYmdLocal();
  elements.dateInput.value = today;
  return today;
}

function ensurePcpEffDateValue() {
  if (elements.pcpEffDate && elements.pcpEffDate.value) {
    return elements.pcpEffDate.value;
  }

  const baseDate = elements.dateInput.value || getTodayYmdLocal();
  if (elements.pcpEffDate) {
    elements.pcpEffDate.value = baseDate;
  }
  return baseDate;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function overlapsMinutes(startA, endA, startB, endB) {
  return Math.max(0, Math.min(endA, endB) - Math.max(startA, startB));
}

function getShiftInfo(dateValue) {
  const selectedDate = dateValue ? new Date(`${dateValue}T00:00:00`) : new Date();
  const weekday = selectedDate.getDay();
  const isFriday = weekday === 5;
  const isWorkday = weekday >= 1 && weekday <= 5;

  return {
    isWorkday,
    startMinutes: 7 * 60,
    endMinutes: isFriday ? 16 * 60 : 17 * 60,
    lunchStart: 12 * 60,
    lunchEnd: 13 * 60
  };
}

function getReferenceMinutes(dateValue, hourOverride) {
  const selectedDate = dateValue ? new Date(`${dateValue}T00:00:00`) : new Date();
  const today = new Date();
  const selectedYmd = selectedDate.toISOString().slice(0, 10);
  const todayYmd = today.toISOString().slice(0, 10);

  if (hourOverride !== null) {
    return hourOverride * 60;
  }

  if (selectedYmd < todayYmd) {
    return 24 * 60;
  }

  if (selectedYmd > todayYmd) {
    return 0;
  }

  return today.getHours() * 60 + today.getMinutes();
}

function calculateDynamicKpis(baseKpis) {
  const shift = getShiftInfo(elements.dateInput.value);
  if (!shift.isWorkday) {
    return null;
  }

  const rawHour = (elements.dynamicHourInput.value || '').trim();
  const parsedHour = rawHour === '' ? null : Number(rawHour);
  const hourOverride = Number.isFinite(parsedHour) ? clamp(parsedHour, 0, 23) : null;
  const referenceMinutes = getReferenceMinutes(elements.dateInput.value, hourOverride);

  const clampedWithinShift = clamp(referenceMinutes, shift.startMinutes, shift.endMinutes);
  let workedMinutes = clampedWithinShift - shift.startMinutes;
  workedMinutes -= overlapsMinutes(shift.startMinutes, clampedWithinShift, shift.lunchStart, shift.lunchEnd);
  workedMinutes = Math.max(0, workedMinutes);

  const totalShiftMinutes = (shift.endMinutes - shift.startMinutes) - (shift.lunchEnd - shift.lunchStart);
  const progress = totalShiftMinutes > 0 ? workedMinutes / totalShiftMinutes : 0;

  const ordensNecessarias = baseKpis.totalOrdens * progress;
  const pecasNecessarias = baseKpis.totalPecasPlanejadas * progress;
  const ordensAtualPercent = baseKpis.totalOrdens > 0 ? (baseKpis.ordensFeitas / baseKpis.totalOrdens) * 100 : 0;
  const pecasAtualPercent = baseKpis.totalPecasPlanejadas > 0
    ? (baseKpis.totalPecasProduzidas / baseKpis.totalPecasPlanejadas) * 100
    : 0;
  const necessarioPercent = progress * 100;

  const eficienciaOrdensDinamica = necessarioPercent > 0
    ? (ordensAtualPercent / necessarioPercent) * 100
    : 100;

  const eficienciaPecasDinamica = necessarioPercent > 0
    ? (pecasAtualPercent / necessarioPercent) * 100
    : 100;

  return {
    ordensNecessarias,
    pecasNecessarias,
    ordensAtualPercent,
    pecasAtualPercent,
    necessarioPercent,
    eficienciaOrdensDinamica,
    eficienciaPecasDinamica,
    progressPercent: necessarioPercent
  };
}

function createBatteryChart(actualPercent, plannedPercent) {
  const wrapper = document.createElement('div');
  wrapper.className = 'battery-wrap';

  const shell = document.createElement('div');
  shell.className = 'battery-shell';

  const fill = document.createElement('div');
  fill.className = 'battery-fill';
  fill.style.width = `${clamp(actualPercent, 0, 100)}%`;

  shell.appendChild(fill);

  if (plannedPercent !== null && plannedPercent !== undefined) {
    const marker = document.createElement('div');
    marker.className = 'battery-marker';
    marker.style.left = `${clamp(plannedPercent, 0, 100)}%`;
    shell.appendChild(marker);
  }

  const cap = document.createElement('div');
  cap.className = 'battery-cap';

  wrapper.appendChild(shell);
  wrapper.appendChild(cap);
  return wrapper;
}

function renderKpis(kpis) {
  elements.kpiSection.innerHTML = '';
  if (!kpis) {
    elements.kpiSection.style.display = 'none';
    return;
  }

  elements.kpiSection.style.display = 'grid';
  const dynamicData = state.dynamicEfficiencyEnabled && isMachineList() ? calculateDynamicKpis(kpis) : null;
  const isMasterKpi = kpis.mode === 'mestra';
  const isProjectKpi = kpis.mode === 'projeto';
  const isAcabamentoKpi = kpis.mode === 'acabamento';
  const isMoldagemKpi = kpis.mode === 'moldagem';

  const ordensLabel = dynamicData ? 'Eficiencia Dinamica de Ordens' : 'Eficiencia de Ordens';
  const pecasLabel = dynamicData ? 'Eficiencia Dinamica de Pecas' : 'Eficiencia de Pecas';
  const ordensValue = kpis.eficienciaOrdens;
  const pecasValue = kpis.eficienciaPecas;
  const ordensActualPercent = kpis.eficienciaOrdens;
  const pecasActualPercent = kpis.eficienciaPecas;
  const ordensPlannedPercent = dynamicData ? dynamicData.progressPercent : null;
  const pecasPlannedPercent = dynamicData ? dynamicData.progressPercent : null;

  const cards = isMasterKpi ? [
    { label: 'Total de Ordens', value: formatInteger(kpis.totalOrdens), layoutClass: 'kpi-top-1' },
    { label: 'Ordens Prontas', value: formatInteger(kpis.ordensProntas), layoutClass: 'kpi-top-2' },
    { label: 'Total de Pecas', value: formatInteger(kpis.totalPecasPlanejadas), layoutClass: 'kpi-top-3' },
    { label: 'Pecas Prontas', value: formatInteger(kpis.totalPecasProntas), layoutClass: 'kpi-top-4' },
    {
      label: '% Teflon em Ordens',
      value: formatPercent(kpis.percentualTeflonOrdens),
      layoutClass: 'kpi-top-5',
      subvalue: `Ordens teflon: ${formatInteger(kpis.ordensTeflon)}`
    },
    {
      label: '% Teflon em Pecas',
      value: formatPercent(kpis.percentualTeflonPecas),
      layoutClass: 'kpi-top-6',
      subvalue: `Pecas teflon: ${formatInteger(kpis.totalPecasTeflon)}`
    },
    {
      label: 'Eficiencia de Ordens',
      value: formatPercent(kpis.eficienciaOrdens),
      layoutClass: 'kpi-top-7',
      batteryActual: kpis.eficienciaOrdens,
      batteryPlanned: null
    },
    {
      label: 'Eficiencia de Pecas',
      value: formatPercent(kpis.eficienciaPecas),
      layoutClass: 'kpi-top-8',
      batteryActual: kpis.eficienciaPecas,
      batteryPlanned: null
    },
    {
      label: 'Ordens Criticas',
      value: formatInteger(kpis.ordensCriticas),
      layoutClass: 'kpi-top-9'
    },
    {
      label: '% Ordens Criticas',
      value: formatPercent(kpis.percentualOrdensCriticas),
      layoutClass: 'kpi-top-10'
    },
    {
      label: 'Pecas Criticas',
      value: formatInteger(kpis.pecasCriticas),
      layoutClass: 'kpi-top-11'
    },
    {
      label: '% Pecas Criticas',
      value: formatPercent(kpis.percentualPecasCriticas),
      layoutClass: 'kpi-top-12'
    }
  ] : isProjectKpi ? [
    { label: 'Total de Programas', value: formatInteger(kpis.totalProgramas), layoutClass: 'kpi-top-1' },
    { label: 'Total de Programas Novos', value: formatInteger(kpis.totalProgramasNovos), layoutClass: 'kpi-top-2' },
    { label: 'Programas Otimizados', value: formatInteger(kpis.programasOtimizados), layoutClass: 'kpi-top-3' },
    { label: 'Projeto Liberado', value: formatInteger(kpis.itensProjetoLiberado), layoutClass: 'kpi-top-4' },
    { label: 'Faltam Liberar', value: formatInteger(kpis.itensFaltandoLiberar), layoutClass: 'kpi-project-missing' },
    {
      label: 'Eficiencia de Liberacao',
      value: formatPercent(kpis.eficienciaLiberacao),
      layoutClass: 'kpi-project-efficiency',
      subvalue: `Liberados: ${formatInteger(kpis.itensProjetoLiberado)} | Faltam: ${formatInteger(kpis.itensFaltandoLiberar)}`,
      batteryActual: kpis.eficienciaLiberacao,
      batteryPlanned: null
    }
  ] : isMoldagemKpi ? [
    {
      label: 'Peso processado total',
      value: formatWeightKg(kpis.pesoProcessadoTotal),
      subvalue: 'Referente a peso em Kg',
      layoutClass: 'kpi-top-1'
    },
    {
      label: 'Total de ordens processadas',
      value: formatInteger(kpis.totalOrdensProcessadas),
      layoutClass: 'kpi-top-2'
    },
    {
      label: 'Total de buchas processadas',
      value: formatInteger(kpis.totalBuchasProcessadas),
      layoutClass: 'kpi-top-3'
    },
    {
      label: 'Eficiencia por Operador',
      value: `${formatInteger((kpis.operadores || []).length)} operador(es) com apontamento`,
      layoutClass: 'kpi-operator-breakdown',
      details: (kpis.operadores || []).map((item) => {
        return `${item.operador} - ${formatInteger(item.ordens)} ordens (${formatPercent(item.percentualOrdens)}) | `
          + `${formatInteger(item.buchas)} buchas (${formatPercent(item.percentualBuchas)}) | `
          + `${formatWeightKg(item.peso)} (${formatPercent(item.percentualPeso)})`;
      })
    }
  ] : [
    {
      label: 'Total de Ordens',
      value: formatInteger(kpis.totalOrdens),
      layoutClass: 'kpi-top-1'
    },
    {
      label: isAcabamentoKpi ? 'Ordens Prontas' : 'Ordens Feitas',
      value: formatInteger(isAcabamentoKpi ? kpis.ordensProntas : kpis.ordensFeitas),
      layoutClass: 'kpi-top-2'
    },
    {
      label: isAcabamentoKpi ? 'Total de Pecas' : 'Pecas Planejadas',
      value: formatInteger(isAcabamentoKpi ? kpis.totalPecas : kpis.totalPecasPlanejadas),
      layoutClass: 'kpi-top-3'
    },
    {
      label: isAcabamentoKpi ? 'Pecas Feitas' : 'Pecas Produzidas',
      value: formatInteger(isAcabamentoKpi ? kpis.pecasFeitas : kpis.totalPecasProduzidas),
      layoutClass: 'kpi-top-4'
    },
    ...(isAcabamentoKpi ? [
      {
        label: 'Ordens Criticas',
        value: formatInteger(kpis.ordensCriticas),
        layoutClass: 'kpi-top-5'
      },
      {
        label: '% Ordens Criticas',
        value: formatPercent(kpis.percentualOrdensCriticas),
        layoutClass: 'kpi-top-6'
      },
      {
        label: 'Pecas Criticas',
        value: formatInteger(kpis.pecasCriticas),
        layoutClass: 'kpi-top-7'
      },
      {
        label: '% Pecas Criticas',
        value: formatPercent(kpis.percentualPecasCriticas),
        layoutClass: 'kpi-top-8'
      },
      {
        label: 'Segregacao por Operador',
        value: `${formatInteger((kpis.operadores || []).length)} operador(es) com apontamento`,
        layoutClass: 'kpi-operator-breakdown',
        details: (kpis.operadores || []).map((item) => {
          return `${item.operador} - ${formatInteger(item.ordens)} ordens (${formatPercent(item.percentualOrdens)}) | `
            + `${formatInteger(item.pecas)} pecas (${formatPercent(item.percentualPecas)}) | `
            + `Criticas: ${formatInteger(item.ordensCriticas)} ordens e ${formatInteger(item.pecasCriticas)} pecas`;
        })
      }
    ] : []),
    {
      label: ordensLabel,
      value: formatPercent(ordensValue),
      layoutClass: 'kpi-bottom-left',
      subvalue: dynamicData
        ? `Necessario: ${formatInteger(dynamicData.ordensNecessarias)} (${formatPercent(dynamicData.necessarioPercent)}) | Realizado: ${formatPercent(dynamicData.ordensAtualPercent)}`
        : '',
      batteryActual: ordensActualPercent,
      batteryPlanned: ordensPlannedPercent
    },
    {
      label: pecasLabel,
      value: formatPercent(pecasValue),
      layoutClass: 'kpi-bottom-right',
      subvalue: dynamicData
        ? `Necessario: ${formatInteger(dynamicData.pecasNecessarias)} (${formatPercent(dynamicData.necessarioPercent)}) | Realizado: ${formatPercent(dynamicData.pecasAtualPercent)}`
        : '',
      batteryActual: pecasActualPercent,
      batteryPlanned: pecasPlannedPercent
    }
  ];

  cards.forEach((item) => {
    const article = document.createElement('article');
    article.className = `kpi-card ${item.layoutClass}`;

    const label = document.createElement('p');
    label.className = 'kpi-label';
    label.textContent = item.label;

    const value = document.createElement('strong');
    value.className = 'kpi-value';
    value.textContent = item.value;

    article.appendChild(label);
    article.appendChild(value);

    if (item.subvalue) {
      const sub = document.createElement('span');
      sub.className = 'kpi-subvalue';
      sub.textContent = item.subvalue;
      article.appendChild(sub);
    }

    if (Array.isArray(item.details) && item.details.length > 0) {
      const detailsList = document.createElement('ul');
      detailsList.className = 'kpi-details-list';
      item.details.forEach((detail) => {
        const detailItem = document.createElement('li');
        detailItem.textContent = detail;
        detailsList.appendChild(detailItem);
      });
      article.appendChild(detailsList);
    }

    if (typeof item.batteryActual === 'number') {
      article.appendChild(createBatteryChart(item.batteryActual, item.batteryPlanned));
    }

    elements.kpiSection.appendChild(article);
  });
}

function isPositiveNumericValue(value) {
  const raw = String(value || '').trim();
  if (raw === '') {
    return false;
  }

  let normalized = raw.replace(/^'+/, '').replace(/\s+/g, '');
  const hasComma = normalized.includes(',');
  const hasDot = normalized.includes('.');

  if (hasComma && hasDot) {
    normalized = normalized.replace(/\./g, '').replace(',', '.');
  } else if (hasComma) {
    normalized = normalized.replace(',', '.');
  }

  const numberValue = Number(normalized);
  return Number.isFinite(numberValue) && numberValue > 0;
}

function populateFilterColumns(columns) {
  elements.filterColumnSelect.innerHTML = '<option value="">Selecione uma coluna</option>';

  getVisibleColumns(columns).forEach((col) => {
    const option = document.createElement('option');
    option.value = col;
    option.textContent = col;
    elements.filterColumnSelect.appendChild(option);
  });
}

function renderTable(columns, rows) {
  elements.tableHead.innerHTML = '';
  elements.tableBody.innerHTML = '';

  if (!rows || rows.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = columns.length || 1;
    td.textContent = 'Nenhuma linha encontrada nesta lista.';
    tr.appendChild(td);
    elements.tableBody.appendChild(tr);
    return;
  }

  const headerRow = document.createElement('tr');
  columns.forEach((col) => {
    const th = document.createElement('th');
    th.textContent = col;
    if (shouldCenterColumn(col)) {
      th.classList.add('center-cell');
    }
    headerRow.appendChild(th);
  });
  elements.tableHead.appendChild(headerRow);

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    if (row.__rowClass) {
      tr.classList.add(row.__rowClass);
    } else if (row.__isQtdDivider) {
      tr.classList.add('qtd-divider');
    } else if (row.__highlight) {
      tr.classList.add('highlight');
    }

    columns.forEach((col) => {
      const td = document.createElement('td');
      td.textContent = row[col] || '';
      const cellClass = row.__cellClasses && row.__cellClasses[col];
      if (cellClass) {
        td.classList.add(cellClass);
      }
      if (shouldCenterColumn(col)) {
        td.classList.add('center-cell');
      }
      tr.appendChild(td);
    });

    elements.tableBody.appendChild(tr);
  });
}

function applyCurrentFilter() {
  const visibleColumns = getVisibleColumns(state.loadedColumns);
  const selectedColumn = elements.filterColumnSelect.value;
  const filterValue = elements.filterValueInput.value || '';
  const normalizedFilter = normalizeForFilter(filterValue);
  const buchaNotEmpty = !isMoldagem() && elements.filterBuchaNotEmpty.checked;
  const missingOrdersOnly = elements.filterProjectMissing.checked;
  const criticalOrdersOnly = elements.filterAcabamentoCritical.checked;
  const duplicatedReqs = buildDuplicatedReqSet(state.loadedRows);

  const filteredRows = state.loadedRows.filter((row) => {
    if (row.__isQtdDivider) {
      return true;
    }

    if (buchaNotEmpty) {
      const buchaValue = row.BUCHA || row['BUCHA PRENSADAS'] || row.__buchaValue || '';
      if (!isPositiveNumericValue(buchaValue)) {
        return false;
      }
    }

    if (missingOrdersOnly) {
      if (!isRowMissingForCurrentList(row)) {
        return false;
      }
    }

    if (criticalOrdersOnly && !isCriticalRowForCurrentList(row, duplicatedReqs)) {
      return false;
    }

    if (!selectedColumn || normalizedFilter === '') {
      return true;
    }

    const normalizedCell = normalizeForFilter(row[selectedColumn]);
    return normalizedCell.includes(normalizedFilter);
  });

  const nonDividerCount = filteredRows.filter((row) => !row.__isQtdDivider).length;
  renderTable(visibleColumns, filteredRows);
  return nonDividerCount;
}

async function init() {
  try {
    const settings = await window.api.getSettings();
    state.settings = settings || {};
    const needsFirstSetup = !String(state.settings.rootFolder || '').trim() || !String(state.settings.fiscalRoot || '').trim();
    state.needsFirstSetup = needsFirstSetup;
    if (needsFirstSetup) {
      state.authUser = { username: '', permissions: { pcp: false, fiscal: false, settings: true, log: true } };
    }
    state.currentListType = elements.listTypeSelect.value || 'acabamento';
    updateMoldagemFilterVisibility();
    updateProjectMissingFilterVisibility();
    updateAcabamentoCriticalFilterVisibility();
    updateDynamicControlsVisibility();
    renderKpis(null);
    renderLogUpdates();
    updateFooterVersion();
    setPcpView('menu');
    setLoginsSectionVisible(false);
    if (elements.pcpEffDate) {
      elements.pcpEffDate.value = ensureDateValue();
    }
    if (elements.pcpDashboardDate) {
      elements.pcpDashboardDate.value = ensureDateValue();
    }
    updatePcpEffTimeButtons();
    renderPcpTable();
    resetPcpEfficiency();
    updateFiscalPendingButtonCount();

    elements.rootFolderInput.value = state.settings.rootFolder || '';
    if (elements.fiscalRootInput) {
      elements.fiscalRootInput.value = state.settings.fiscalRoot || '';
    }
    if (elements.trackingApiKeyInput) {
      elements.trackingApiKeyInput.value = state.settings.trackingApiKey || '';
    }
    applyTheme(state.settings.theme || 'light');
    applyPermissionsToTabs();
    if (needsFirstSetup) {
      switchTab('settings');
      showToast('Configure a Pasta raiz e o Caminho Fiscal (R:) antes do login.');
    } else {
      openAppLoginModal();
    }
  } catch (error) {
    showToast(`Erro ao carregar configuracao: ${error.message}`, true);
  }
}

async function handleSelectFolder() {
  try {
    const selected = await window.api.selectRootFolder();
    if (!selected) {
      return;
    }

    elements.rootFolderInput.value = selected;
    state.settings = {
      ...(state.settings || {}),
      rootFolder: selected
    };
    showToast('Pasta raiz selecionada.');
  } catch (error) {
    showToast(`Erro ao selecionar pasta: ${error.message}`, true);
  }
}

async function handleSelectFiscalRoot() {
  try {
    const selected = await window.api.selectFiscalRootFolder();
    if (!selected) {
      return;
    }

    if (elements.fiscalRootInput) {
      elements.fiscalRootInput.value = selected;
    }
    state.settings = {
      ...(state.settings || {}),
      fiscalRoot: selected
    };
    showToast('Caminho Fiscal selecionado.');

    if (state.isLoginsSectionUnlocked) {
      await loadUsers();
      renderUsersTable();
    }
  } catch (error) {
    showToast(`Erro ao selecionar caminho fiscal: ${error.message}`, true);
  }
}

async function handleSaveSettings() {
  try {
    const nextTheme = elements.body.getAttribute('data-theme') || 'light';
    const saved = await window.api.saveSettings({
      rootFolder: elements.rootFolderInput.value || '',
      fiscalRoot: elements.fiscalRootInput?.value || '',
      trackingApiKey: elements.trackingApiKeyInput?.value || '',
      theme: nextTheme
    });

    state.settings = saved || {};
    showToast('Configuracao salva.');

    await loadUsers();
    renderUsersTable();

    if (state.needsFirstSetup) {
      state.needsFirstSetup = false;
      state.authUser = null;
      applyPermissionsToTabs();
      openAppLoginModal();
    }
  } catch (error) {
    showToast(`Erro ao salvar configuracao: ${error.message}`, true);
  }
}

async function handleFiscalTrackingLookup() {
  if (!elements.fiscalTrackingCode || !elements.fiscalTrackingResult) {
    return;
  }

  const rawCode = String(elements.fiscalTrackingCode.value || '').trim();
  if (!rawCode) {
    elements.fiscalTrackingResult.classList.add('error');
    elements.fiscalTrackingResult.textContent = 'Informe um código de rastreio.';
    return;
  }

  elements.fiscalTrackingResult.classList.remove('error');
  elements.fiscalTrackingResult.textContent = 'Consultando...';

  try {
    const result = await window.api.fiscalTrackingLookup({ code: rawCode });
    if (!result || !result.valid) {
      elements.fiscalTrackingResult.classList.add('error');
      elements.fiscalTrackingResult.textContent = 'Código não encontrado/ inválido.';
      return;
    }

    const parts = [`Status: ${result.status || '-'}`];
    if (result.fiscal && (result.fiscal.cliente || result.fiscal.nf)) {
      const label = [
        result.fiscal.cliente ? `Cliente: ${result.fiscal.cliente}` : '',
        result.fiscal.nf ? `NF: ${result.fiscal.nf}` : ''
      ].filter(Boolean).join(' | ');
      if (label) {
        parts.push(label);
      }
    }
    if (result.destino) {
      parts.push(`Destino: ${result.destino}`);
    }
    if (result.local) {
      parts.push(`Local: ${result.local}`);
    }
    if (result.updatedAt) {
      parts.push(`Atualizado: ${result.updatedAt}`);
    }
    elements.fiscalTrackingResult.textContent = parts.join(' | ');
  } catch (error) {
    elements.fiscalTrackingResult.classList.add('error');
    elements.fiscalTrackingResult.textContent = error.message || 'Erro ao consultar rastreio.';
  }
}

function handleToggleTheme() {
  const next = elements.themeSwitch && elements.themeSwitch.checked ? 'dark' : 'light';
  applyTheme(next);
  showToast(`Tema ${next === 'dark' ? 'escuro' : 'claro'} aplicado.`);
}

async function handleFiscalSearchItems() {
  try {
    const fiscalRoot = String(state.settings?.fiscalRoot || '').trim();
    if (!fiscalRoot) {
      throw new Error('Configure o Caminho Fiscal (R:) na aba Configuracao.');
    }

    const nf = String(elements.fiscalNfNumber?.value || '').trim();
    const client = String(elements.fiscalNfClient?.value || '').trim();
    const identifiers = parseFiscalIdentifiers(elements.fiscalNfIdentifiers?.value || '');

    if (!nf || !client || !identifiers.length) {
      if (elements.fiscalNfError) {
        elements.fiscalNfError.textContent = 'Preencha NF, Cliente e ao menos 1 Pedido/Requisicao.';
      }
      return;
    }

    if (elements.fiscalNfError) {
      elements.fiscalNfError.textContent = '';
    }

    const result = await window.api.fiscalPreviewItems({ identifiers });
    const items = (result && result.items) || [];
    state.fiscalPreviewItems = items;

    if (!items.length) {
      if (elements.fiscalNfError) {
        elements.fiscalNfError.textContent = 'Nenhum item encontrado para os pedidos/requisicoes informados.';
      }
      if (elements.fiscalNfConfirm) {
        elements.fiscalNfConfirm.disabled = true;
      }
      if (elements.fiscalNfPreview) {
        elements.fiscalNfPreview.hidden = true;
      }
      return;
    }

    renderFiscalPreview(items);
    if (elements.fiscalNfConfirm) {
      elements.fiscalNfConfirm.disabled = false;
    }
  } catch (error) {
    if (elements.fiscalNfError) {
      elements.fiscalNfError.textContent = error.message || 'Erro ao buscar itens.';
    }
  }
}

async function handleFiscalConfirmCadastro() {
  try {
    const fiscalRoot = String(state.settings?.fiscalRoot || '').trim();
    if (!fiscalRoot) {
      throw new Error('Configure o Caminho Fiscal (R:) na aba Configuracao.');
    }

    const nf = String(elements.fiscalNfNumber?.value || '').trim();
    const client = String(elements.fiscalNfClient?.value || '').trim();
    const identifiers = parseFiscalIdentifiers(elements.fiscalNfIdentifiers?.value || '');
    const faturadaAt = String(elements.fiscalNfDatetime?.value || '').trim();
    const user = String(state.authUser?.username || state.fiscalUser || '').trim();

    if (!nf || !client || !identifiers.length) {
      throw new Error('Preencha NF, Cliente e ao menos 1 Pedido/Requisicao.');
    }
    if (!faturadaAt) {
      throw new Error('Informe a Data/Horario faturado.');
    }
    if (!user) {
      throw new Error('Usuario nao identificado. Acesse a aba FISCAL novamente.');
    }

    if (!state.fiscalPreviewItems.length) {
      throw new Error('Busque os itens antes de cadastrar.');
    }

    const result = await window.api.fiscalRegisterNf({
      nf,
      client,
      identifiers,
      faturadaAt,
      user
    });

    const inserted = (result && result.inserted) || [];
    renderFiscalTable(inserted);
    closeFiscalNfModal();
    await loadFiscalNfs();
    showToast(`NF cadastrada. ${formatInteger(inserted.length)} item(ns) gravado(s).`);
  } catch (error) {
    if (elements.fiscalNfError) {
      elements.fiscalNfError.textContent = error.message || 'Erro ao cadastrar NF.';
    }
  }
}

function parseDatetimeLocalToDate(value) {
  const text = String(value || '').trim();
  if (!text) {
    return null;
  }
  const dt = new Date(text);
  if (Number.isNaN(dt.getTime())) {
    return null;
  }
  return dt;
}

async function confirmAndSaveFiscalEdit() {
  const nf = state.fiscalSelectionNf;
  if (!nf) {
    return;
  }
  const editorUser = String(state.authUser?.username || state.fiscalUser || '').trim();
  if (!editorUser) {
    if (elements.fiscalConfirmError) {
      elements.fiscalConfirmError.textContent = 'Usuário logado não identificado.';
    }
    return;
  }

  try {
    const mode = state.fiscalConfirmMode || 'edit';
    let updated = 0;

    if (mode === 'delete') {
      const reason = String(state.fiscalDeleteReason || '').trim();
      const result = await window.api.fiscalDeleteNf({
        nf,
        reason,
        editorUser,
        editorAt: new Date().toISOString()
      });
      updated = (result && result.deletedRows) || 0;
    } else {
      const status = String(elements.fiscalEditStatus?.value || 'Faturado').trim();
      const rastreio = String(elements.fiscalEditRastreio?.value || '').trim();
      const despacheDt = parseDatetimeLocalToDate(elements.fiscalEditDespache?.value || '');
      const dataDespache = despacheDt ? despacheDt.toLocaleDateString('pt-BR') : '';

      const result = await window.api.fiscalUpdateNf({
        nf,
        status,
        rastreio,
        dataDespache,
        editorUser,
        editorAt: new Date().toISOString()
      });

      updated = (result && result.updated) || 0;
    }

    closeFiscalConfirmModal();
    closeFiscalEditModal();
    closeFiscalDeleteModal();
    await loadFiscalNfs();

    if (mode === 'delete') {
      state.fiscalSelectionNf = null;
      state.fiscalNfItems = [];
      renderFiscalScreen();
      showToast(`NF apagada. ${formatInteger(updated)} linha(s) removida(s).`);
    } else {
      await openFiscalNfDetails(nf);
      showToast(`NF atualizada. ${formatInteger(updated)} linha(s) alterada(s).`);
    }
  } catch (error) {
    if (elements.fiscalConfirmError) {
      elements.fiscalConfirmError.textContent = error.message || 'Erro ao atualizar NF.';
    }
  }
}

async function handleLoadList() {
  try {
    const date = ensureDateValue();
    const listType = elements.listTypeSelect.value;
    state.currentListType = listType;
    updateMoldagemFilterVisibility();
    updateProjectMissingFilterVisibility();
    updateAcabamentoCriticalFilterVisibility();
    updateDynamicControlsVisibility();

    const data = await window.api.loadList({ date, listType });
    state.loadedColumns = data.columns || [];
    state.loadedRows = data.rows || [];
    state.loadedKpis = data.kpis || null;
    populateFilterColumns(state.loadedColumns);
    elements.filterColumnSelect.value = '';
    elements.filterValueInput.value = '';
    elements.filterBuchaNotEmpty.checked = false;
    elements.filterProjectMissing.checked = false;
    elements.filterAcabamentoCritical.checked = false;

    elements.resultPath.textContent = `${data.filePath} | Aba: ${data.sheetName} | Linhas: ${data.rows.length}`;
    renderKpis(state.loadedKpis);
    renderTable(getVisibleColumns(state.loadedColumns), state.loadedRows);
    renderPcpTable();
    showToast('Lista carregada com sucesso.');
  } catch (error) {
    state.loadedColumns = [];
    state.loadedRows = [];
    state.loadedKpis = null;
    populateFilterColumns([]);
    elements.filterColumnSelect.value = '';
    elements.filterValueInput.value = '';
    elements.filterBuchaNotEmpty.checked = false;
    elements.filterProjectMissing.checked = false;
    elements.filterAcabamentoCritical.checked = false;
    elements.resultPath.textContent = 'Nenhum arquivo carregado.';
    renderKpis(null);
    renderTable([], []);
    renderPcpTable();
    showToast(error.message || 'Erro ao carregar lista.', true);
  }
}

function handleApplyFilter() {
  if (!state.loadedRows.length) {
    showToast('Carregue uma lista antes de filtrar.', true);
    return;
  }

  const hasColumnFilter = !!elements.filterColumnSelect.value && !!elements.filterValueInput.value.trim();
  const hasBuchaFilter = !isMoldagem() && elements.filterBuchaNotEmpty.checked;
  const hasProjectMissingFilter = elements.filterProjectMissing.checked;
  const hasAcabamentoCriticalFilter = elements.filterAcabamentoCritical.checked;

  if (!hasColumnFilter && !hasBuchaFilter && !hasProjectMissingFilter && !hasAcabamentoCriticalFilter) {
    const message = isMoldagem()
      ? 'Informe um filtro por coluna/valor.'
      : 'Informe um filtro por coluna/valor, marque Ordens faltantes ou Ordens criticas.';
    showToast(message, true);
    return;
  }

  const count = applyCurrentFilter();
  showToast(`Filtro aplicado. ${count} linha(s) encontrada(s).`);
}

function handleClearFilter() {
  elements.filterColumnSelect.value = '';
  elements.filterValueInput.value = '';
  elements.filterBuchaNotEmpty.checked = false;
  elements.filterProjectMissing.checked = false;
  elements.filterAcabamentoCritical.checked = false;
  renderTable(getVisibleColumns(state.loadedColumns), state.loadedRows);
  showToast('Filtro limpo.');
}

function handleToggleDynamicEfficiency() {
  if (!isMachineList()) {
    return;
  }

  state.dynamicEfficiencyEnabled = !!elements.dynamicEffSwitch.checked;
  renderKpis(state.loadedKpis);
}

function openAreaFromMenu(area) {
  if (area === 'main') {
    switchTab('main');
    return;
  }
  if (area === 'pcp') {
    requestPcpAccess();
    return;
  }
  if (area === 'fiscal') {
    void requestFiscalAccess();
    return;
  }
  if (area === 'rh') {
    requestRhAccess();
    return;
  }
  if (area === 'settings') {
    if (!canAccessArea('settings')) {
      showToast('Usuário sem permissão para acessar Configuração.', true);
      return;
    }
    switchTab('settings');
    return;
  }
  if (area === 'log') {
    if (!canAccessArea('log')) {
      showToast('Usuário sem permissão para acessar LOG.', true);
      return;
    }
    switchTab('log');
  }
}

async function handleRunBackupNow() {
  const btn = elements.btnRunBackupNow || null;
  const label = btn ? btn.querySelector('.backup-now-label') : null;
  if (btn && btn.disabled) {
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.classList.add('is-loading');
  }
  if (label) {
    label.textContent = 'Executando...';
  }

  try {
    const result = await window.api.runBackupNow();
    const folderPath = String(result?.folderPath || '').trim();
    if (folderPath) {
      showToast(`Backup concluído: ${folderPath}`);
      return;
    }
    showToast('Backup concluído.');
  } catch (error) {
    showToast(`Erro ao executar backup: ${error.message}`, true);
  } finally {
    if (label) {
      label.textContent = 'Executar backup agora';
    }
    if (btn) {
      btn.disabled = false;
      btn.classList.remove('is-loading');
    }
  }
}

async function handleRestoreBackupConfirm() {
  const backupFolderName = String(elements.backupRestoreSelect?.value || '').trim();
  if (!backupFolderName) {
    if (elements.backupRestoreError) {
      elements.backupRestoreError.textContent = 'Selecione um backup para restaurar.';
    }
    return;
  }
  const pass1 = String(elements.backupRestorePass1?.value || '').trim();
  const pass2 = String(elements.backupRestorePass2?.value || '').trim();
  const pass3 = String(elements.backupRestorePass3?.value || '').trim();
  if (!pass1 || !pass2 || !pass3) {
    if (elements.backupRestoreError) {
      elements.backupRestoreError.textContent = 'Informe as 3 confirmações de senha.';
    }
    return;
  }

  const confirmBtn = elements.backupRestoreConfirm || null;
  if (confirmBtn) {
    confirmBtn.disabled = true;
  }
  if (elements.backupRestoreError) {
    elements.backupRestoreError.textContent = '';
  }

  try {
    const result = await window.api.backupRestore({
      backupFolderName,
      adminPasswords: [pass1, pass2, pass3]
    });
    const restoredCount = Number(result?.restoredCount || 0);
    closeBackupRestoreModal();
    showToast(`Backup restaurado com sucesso (${restoredCount} arquivo(s)). Reinicie o app para aplicar tudo.`);
  } catch (error) {
    if (elements.backupRestoreError) {
      elements.backupRestoreError.textContent = error.message || 'Erro ao restaurar backup.';
    }
  } finally {
    if (confirmBtn) {
      confirmBtn.disabled = false;
    }
  }
}

function isTypingContextTarget(target) {
  if (!target) {
    return false;
  }
  const tag = String(target.tagName || '').toUpperCase();
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') {
    return true;
  }
  return !!target.isContentEditable;
}

function handleMenuShortcutKeydown(event) {
  if (!event || event.defaultPrevented) {
    return;
  }
  if (event.ctrlKey || event.altKey || event.metaKey) {
    return;
  }
  if (!elements.panelMenu || !elements.panelMenu.classList.contains('active')) {
    return;
  }
  if (elements.appLoginModal && !elements.appLoginModal.hidden) {
    return;
  }
  if (isTypingContextTarget(event.target)) {
    return;
  }

  const key = String(event.key || '').trim();
  const code = String(event.code || '').trim();
  const resolved = key || (code.startsWith('Digit') || code.startsWith('Numpad') ? code.replace(/^(Digit|Numpad)/, '') : '');
  const map = {
    '1': 'main',
    '2': 'pcp',
    '3': 'fiscal',
    '4': 'rh',
    '5': 'settings',
    '6': 'log'
  };
  const area = map[resolved];
  if (!area) {
    return;
  }
  event.preventDefault();
  openAreaFromMenu(area);
}

document.addEventListener('keydown', handleMenuShortcutKeydown);

if (elements.tabMain) {
  elements.tabMain.addEventListener('click', () => switchTab('main'));
}
if (elements.tabMenu) {
  elements.tabMenu.addEventListener('click', () => switchTab('menu'));
}
if (elements.tabPcp) {
  elements.tabPcp.addEventListener('click', () => {
    if (elements.tabPcp.classList.contains('active')) {
      return;
    }

    requestPcpAccess();
  });
}
if (elements.tabFiscal) {
  elements.tabFiscal.addEventListener('click', async () => {
    if (elements.tabFiscal.classList.contains('active')) {
      return;
    }

    await requestFiscalAccess();
  });
}
if (elements.tabRh) {
  elements.tabRh.addEventListener('click', () => {
    if (elements.tabRh.classList.contains('active')) {
      return;
    }
    requestRhAccess();
  });
}
if (elements.tabSettings) {
  elements.tabSettings.addEventListener('click', () => {
    if (!canAccessArea('settings')) {
      showToast('Usuário sem permissão para acessar Configuração.', true);
      return;
    }
    switchTab('settings');
  });
}
if (elements.tabLog) {
  elements.tabLog.addEventListener('click', () => {
    if (!canAccessArea('log')) {
      showToast('Usuário sem permissão para acessar LOG.', true);
      return;
    }
    switchTab('log');
  });
}
if (elements.pcpOpenAcomp) {
  elements.pcpOpenAcomp.addEventListener('click', async () => {
    try {
      await openPcpAcompanhamento();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar acompanhamento.', true);
    }
  });
}
if (elements.pcpOpenEff) {
  elements.pcpOpenEff.addEventListener('click', async () => {
    try {
      await openPcpEfficiency();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar eficiência.', true);
    }
  });
}
if (elements.pcpOpenDashboard) {
  elements.pcpOpenDashboard.addEventListener('click', async () => {
    try {
      await openPcpDashboard();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar dashboard diário.', true);
    }
  });
}
if (elements.pcpBackMenu) {
  elements.pcpBackMenu.addEventListener('click', () => {
    setPcpView('menu');
  });
}
if (elements.pcpEffBackMenu) {
  elements.pcpEffBackMenu.addEventListener('click', () => {
    setPcpView('menu');
  });
}
if (elements.pcpDashboardBackMenu) {
  elements.pcpDashboardBackMenu.addEventListener('click', () => {
    setPcpView('menu');
  });
}
if (elements.pcpEffRefresh) {
  elements.pcpEffRefresh.addEventListener('click', async () => {
    try {
      await loadPcpEfficiencySnapshot();
      renderPcpEfficiency();
      showToast('Painel de eficiência atualizado.');
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar eficiência.', true);
    }
  });
}
if (elements.pcpEffExport) {
  elements.pcpEffExport.addEventListener('click', async () => {
    try {
      const filePath = await exportPcpEfficiencyImage();
      if (filePath) {
        showToast('Imagem exportada com sucesso.');
      }
    } catch (error) {
      showToast(error.message || 'Erro ao exportar imagem da eficiência.', true);
    }
  });
}
if (elements.pcpDashboardRefresh) {
  elements.pcpDashboardRefresh.addEventListener('click', async () => {
    try {
      await loadPcpDashboardSnapshot();
      await loadPcpMoldagemSnapshot();
      renderPcpDashboard();
      showToast('Dashboard diário atualizado.');
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar dashboard diário.', true);
    }
  });
}
if (elements.pcpDashboardExportPdf) {
  elements.pcpDashboardExportPdf.addEventListener('click', async () => {
    try {
      const filePath = await exportPcpDashboardPdf();
      if (filePath) {
        showToast('PDF do dashboard exportado com sucesso.');
      }
    } catch (error) {
      showToast(error.message || 'Erro ao exportar PDF do dashboard.', true);
    }
  });
}
if (elements.pcpEffDate) {
  elements.pcpEffDate.addEventListener('change', async () => {
    resetPcpEfficiency();
    if (state.pcpView === 'eff') {
      try {
        await loadPcpEfficiencySnapshot();
        renderPcpEfficiency();
      } catch (error) {
        showToast(error.message || 'Erro ao atualizar data da eficiência.', true);
      }
    }
  });
}
if (elements.pcpDashboardDate) {
  elements.pcpDashboardDate.addEventListener('change', async () => {
    state.pcpDashboardSnapshot = null;
    state.pcpMoldagemSnapshot = null;
    if (state.pcpView === 'dashboard') {
      try {
        await loadPcpDashboardSnapshot();
        await loadPcpMoldagemSnapshot();
        renderPcpDashboard();
      } catch (error) {
        showToast(error.message || 'Erro ao atualizar data do dashboard.', true);
      }
    }
  });
}
[elements.pcpEffTimeAuto, elements.pcpEffTime10, elements.pcpEffTime12, elements.pcpEffTime15]
  .filter(Boolean)
  .forEach((chip) => {
    chip.addEventListener('click', () => {
      state.pcpEfficiencyHourMode = chip.dataset.hourMode || 'auto';
      updatePcpEffTimeButtons();
      renderPcpEfficiency();
    });
  });
elements.btnSelectFolder.addEventListener('click', handleSelectFolder);
if (elements.btnSelectFiscalRoot) {
  elements.btnSelectFiscalRoot.addEventListener('click', handleSelectFiscalRoot);
}
if (elements.btnRunBackupNow) {
  elements.btnRunBackupNow.addEventListener('click', handleRunBackupNow);
}
if (elements.btnRestoreBackup) {
  elements.btnRestoreBackup.addEventListener('click', openBackupRestoreModal);
}
if (elements.backupRestoreCancel) {
  elements.backupRestoreCancel.addEventListener('click', closeBackupRestoreModal);
}
if (elements.backupRestoreBackdrop) {
  elements.backupRestoreBackdrop.addEventListener('click', closeBackupRestoreModal);
}
if (elements.backupRestoreConfirm) {
  elements.backupRestoreConfirm.addEventListener('click', handleRestoreBackupConfirm);
}
[elements.backupRestorePass1, elements.backupRestorePass2, elements.backupRestorePass3]
  .filter(Boolean)
  .forEach((input) => {
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        handleRestoreBackupConfirm();
      } else if (event.key === 'Escape') {
        closeBackupRestoreModal();
      }
    });
  });
if (elements.btnUserAdd) {
  elements.btnUserAdd.addEventListener('click', () => {
    requireAdminSettings(() => openUserModal());
  });
}
if (elements.btnUnlockLogins) {
  elements.btnUnlockLogins.addEventListener('click', () => {
    requireAdminSettings(async () => {
      try {
        await loadUsers();
        renderUsersTable();
        setLoginsSectionVisible(true);
        showToast('Seção de logins liberada.');
      } catch (error) {
        showToast(error.message || 'Erro ao carregar logins.', true);
      }
    });
  });
}
elements.btnSaveSettings.addEventListener('click', handleSaveSettings);
elements.themeSwitch.addEventListener('change', handleToggleTheme);
elements.btnLoad.addEventListener('click', handleLoadList);
elements.listTypeSelect.addEventListener('change', () => {
  state.currentListType = elements.listTypeSelect.value || 'acabamento';
  resetPcpSummary();
  updateMoldagemFilterVisibility();
  updateProjectMissingFilterVisibility();
  updateAcabamentoCriticalFilterVisibility();
  updateDynamicControlsVisibility();
  renderKpis(isMachineList() ? state.loadedKpis : null);
});
elements.dateInput.addEventListener('change', () => {
  resetPcpSummary();
  if (elements.pcpEffDate && elements.dateInput.value) {
    elements.pcpEffDate.value = elements.dateInput.value;
  }
  if (elements.pcpDashboardDate && elements.dateInput.value) {
    elements.pcpDashboardDate.value = elements.dateInput.value;
  }
  resetPcpEfficiency();
  state.pcpDashboardSnapshot = null;
});
elements.btnApplyFilter.addEventListener('click', handleApplyFilter);
elements.btnClearFilter.addEventListener('click', handleClearFilter);
elements.filterBuchaNotEmpty.addEventListener('change', () => {
  if (!state.loadedRows.length) {
    return;
  }

  if (isMoldagem()) {
    return;
  }

  const count = applyCurrentFilter();
  showToast(`Filtro Ordens de moldagem aplicado. ${count} linha(s) encontrada(s).`);
});
elements.filterProjectMissing.addEventListener('change', () => {
  if (!state.loadedRows.length) {
    return;
  }

  const count = applyCurrentFilter();
  showToast(`Filtro Ordens faltantes aplicado. ${count} linha(s) encontrada(s).`);
});
elements.filterAcabamentoCritical.addEventListener('change', () => {
  if (!state.loadedRows.length) {
    return;
  }

  const count = applyCurrentFilter();
  showToast(`Filtro Ordens criticas aplicado. ${count} linha(s) encontrada(s).`);
});
elements.filterValueInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    handleApplyFilter();
  }
});
elements.dynamicEffSwitch.addEventListener('change', handleToggleDynamicEfficiency);
elements.dynamicHourInput.addEventListener('input', () => {
  if (!state.dynamicEfficiencyEnabled) {
    return;
  }

  renderKpis(state.loadedKpis);
});

if (elements.appLoginConfirm) {
  elements.appLoginConfirm.addEventListener('click', handleAppLoginConfirm);
}
if (elements.appLoginConfig) {
  elements.appLoginConfig.addEventListener('click', handleAppLoginOpenSettings);
}
if (elements.appLoginFiscalRoot) {
  elements.appLoginFiscalRoot.addEventListener('click', async () => {
    await handleSelectFiscalRoot();
    if (elements.appLoginError) {
      elements.appLoginError.textContent = '';
    }
    elements.appLoginUsername?.focus?.();
  });
}
if (elements.appLoginUsername || elements.appLoginPassword) {
  [elements.appLoginUsername, elements.appLoginPassword].filter(Boolean).forEach((input) => {
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        handleAppLoginConfirm();
      }
    });
  });
}

if (elements.fiscalConfirmOk) {
  elements.fiscalConfirmOk.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      confirmAndSaveFiscalEdit();
    } else if (event.key === 'Escape') {
      closeFiscalConfirmModal();
    }
  });
}

if (elements.fiscalOpenNfs) {
  elements.fiscalOpenNfs.addEventListener('click', async () => {
    try {
      await loadFiscalNfs();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar NFs do banco.', true);
    }
    state.fiscalSelectionNf = null;
    state.fiscalNfItems = [];
    state.fiscalView = 'nfs';
    renderFiscalScreen();
  });
}
if (elements.menuOpenMain) {
  elements.menuOpenMain.addEventListener('click', () => openAreaFromMenu('main'));
}
if (elements.menuOpenPcp) {
  elements.menuOpenPcp.addEventListener('click', () => openAreaFromMenu('pcp'));
}
if (elements.menuOpenFiscal) {
  elements.menuOpenFiscal.addEventListener('click', () => openAreaFromMenu('fiscal'));
}
if (elements.menuOpenRh) {
  elements.menuOpenRh.addEventListener('click', () => openAreaFromMenu('rh'));
}
if (elements.menuOpenSettings) {
  elements.menuOpenSettings.addEventListener('click', () => openAreaFromMenu('settings'));
}
if (elements.menuOpenLog) {
  elements.menuOpenLog.addEventListener('click', () => openAreaFromMenu('log'));
}

if (elements.fiscalOpenCompras) {
  elements.fiscalOpenCompras.addEventListener('click', async () => {
    try {
      await openFiscalComprasView();
    } catch (error) {
      showToast(error.message || 'Erro ao abrir Solicitação de Compras.', true);
    }
  });
}

if (elements.fiscalBackMenu) {
  elements.fiscalBackMenu.addEventListener('click', () => {
    state.fiscalSelectionNf = null;
    state.fiscalNfItems = [];
    state.fiscalView = 'menu';
    renderFiscalScreen();
  });
}

if (elements.fiscalComprasBackMenu) {
  elements.fiscalComprasBackMenu.addEventListener('click', () => {
    state.fiscalView = 'menu';
    renderFiscalScreen();
  });
}

if (elements.fiscalBtnCadastrarNf) {
  elements.fiscalBtnCadastrarNf.addEventListener('click', () => {
    openFiscalNfModal();
  });
}
if (elements.fiscalTrackingCheck) {
  elements.fiscalTrackingCheck.addEventListener('click', handleFiscalTrackingLookup);
}
if (elements.fiscalTrackingCode) {
  elements.fiscalTrackingCode.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      handleFiscalTrackingLookup();
    }
  });
}
if (elements.fiscalBtnPesquisarNf) {
  elements.fiscalBtnPesquisarNf.addEventListener('click', () => {
    openFiscalSearchModal();
  });
}
if (elements.fiscalBtnHistorico) {
  elements.fiscalBtnHistorico.addEventListener('click', async () => {
    try {
      await loadFiscalHistory();
      state.fiscalSelectionNf = null;
      state.fiscalNfItems = [];
      state.fiscalView = 'history';
      renderFiscalScreen();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar historico.', true);
    }
  });
}
if (elements.fiscalBtnPendencias) {
  elements.fiscalBtnPendencias.addEventListener('click', async () => {
    try {
      await loadFiscalNfs();
      state.fiscalSelectionNf = null;
      state.fiscalNfItems = [];
      state.fiscalView = 'pending';
      renderFiscalScreen();
    } catch (error) {
      showToast(error.message || 'Erro ao carregar pendências fiscais.', true);
    }
  });
}
if (elements.fiscalBtnVoltar) {
  elements.fiscalBtnVoltar.addEventListener('click', () => {
    state.fiscalSelectionNf = null;
    state.fiscalNfItems = [];
    state.fiscalView = 'nfs';
    renderFiscalScreen();
  });
}
if (elements.fiscalBtnEditarNf) {
  elements.fiscalBtnEditarNf.addEventListener('click', () => {
    openFiscalEditModal();
  });
}
if (elements.fiscalBtnApagarNf) {
  elements.fiscalBtnApagarNf.addEventListener('click', () => {
    openFiscalDeleteModal();
  });
}

if (elements.fiscalNfCancel) {
  elements.fiscalNfCancel.addEventListener('click', () => closeFiscalNfModal());
}
if (elements.fiscalNfBackdrop) {
  elements.fiscalNfBackdrop.addEventListener('click', () => closeFiscalNfModal());
}
if (elements.fiscalNfSearch) {
  elements.fiscalNfSearch.addEventListener('click', handleFiscalSearchItems);
}
if (elements.fiscalNfConfirm) {
  elements.fiscalNfConfirm.addEventListener('click', handleFiscalConfirmCadastro);
}
if (elements.fiscalNfClient) {
  elements.fiscalNfClient.addEventListener('input', () => {
    renderFiscalClientSuggestions(elements.fiscalNfClient.value || '');
  });
  elements.fiscalNfClient.addEventListener('focus', () => {
    renderFiscalClientSuggestions(elements.fiscalNfClient.value || '');
  });
  elements.fiscalNfClient.addEventListener('blur', () => {
    setTimeout(() => {
      hideFiscalClientSuggestions();
    }, 120);
  });
}

if (elements.fiscalEditCancel) {
  elements.fiscalEditCancel.addEventListener('click', () => closeFiscalEditModal());
}
if (elements.fiscalEditBackdrop) {
  elements.fiscalEditBackdrop.addEventListener('click', () => closeFiscalEditModal());
}
if (elements.fiscalEditConfirm) {
  elements.fiscalEditConfirm.addEventListener('click', () => {
    if (elements.fiscalEditError) {
      elements.fiscalEditError.textContent = '';
    }
    state.fiscalConfirmMode = 'edit';
    closeFiscalEditModal();
    openFiscalConfirmModal();
  });
}

if (elements.fiscalConfirmCancel) {
  elements.fiscalConfirmCancel.addEventListener('click', () => {
    const mode = state.fiscalConfirmMode || 'edit';
    closeFiscalConfirmModal();
    if (mode === 'delete') {
      reopenFiscalDeleteModal();
    } else {
      reopenFiscalEditModal();
    }
  });
}
if (elements.fiscalConfirmBackdrop) {
  elements.fiscalConfirmBackdrop.addEventListener('click', () => {
    const mode = state.fiscalConfirmMode || 'edit';
    closeFiscalConfirmModal();
    if (mode === 'delete') {
      reopenFiscalDeleteModal();
    } else {
      reopenFiscalEditModal();
    }
  });
}
if (elements.fiscalConfirmOk) {
  elements.fiscalConfirmOk.addEventListener('click', confirmAndSaveFiscalEdit);
}

if (elements.fiscalSearchCancel) {
  elements.fiscalSearchCancel.addEventListener('click', () => closeFiscalSearchModal());
}
if (elements.fiscalSearchBackdrop) {
  elements.fiscalSearchBackdrop.addEventListener('click', () => closeFiscalSearchModal());
}
if (elements.fiscalSearchConfirm) {
  elements.fiscalSearchConfirm.addEventListener('click', handleFiscalSearchNf);
}
if (elements.fiscalSearchNf) {
  elements.fiscalSearchNf.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      handleFiscalSearchNf();
    } else if (event.key === 'Escape') {
      closeFiscalSearchModal();
    }
  });
}

if (elements.fiscalDeleteCancel) {
  elements.fiscalDeleteCancel.addEventListener('click', () => closeFiscalDeleteModal());
}
if (elements.fiscalDeleteBackdrop) {
  elements.fiscalDeleteBackdrop.addEventListener('click', () => closeFiscalDeleteModal());
}
if (elements.fiscalDeleteConfirm) {
  elements.fiscalDeleteConfirm.addEventListener('click', () => {
    const reason = String(elements.fiscalDeleteReason?.value || '').trim();
    if (!reason) {
      if (elements.fiscalDeleteError) {
        elements.fiscalDeleteError.textContent = 'Informe o motivo para apagar a NF.';
      }
      return;
    }

    if (elements.fiscalDeleteError) {
      elements.fiscalDeleteError.textContent = '';
    }

    state.fiscalDeleteReason = reason;
    state.fiscalConfirmMode = 'delete';
    closeFiscalDeleteModal();
    openFiscalConfirmModal();
  });
}

if (elements.fiscalComprasRefresh) {
  elements.fiscalComprasRefresh.addEventListener('click', async () => {
    try {
      await loadFiscalComprasContext();
      await loadFiscalComprasRows();
      applyFiscalComprasPermissionsUi();
      showToast('Solicitações de compra atualizadas.');
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar solicitações de compra.', true);
    }
  });
}

if (elements.fiscalComprasNew) {
  elements.fiscalComprasNew.addEventListener('click', () => {
    if (!canFiscalComprasCreate()) {
      showToast('Usuário sem permissão para criar solicitação de compra.', true);
      return;
    }
    resetFiscalComprasForm();
    setFiscalComprasTab('form');
  });
}

if (elements.fiscalComprasTabForm) {
  elements.fiscalComprasTabForm.addEventListener('click', () => setFiscalComprasTab('form'));
}
if (elements.fiscalComprasTabConsulta) {
  elements.fiscalComprasTabConsulta.addEventListener('click', async () => {
    setFiscalComprasTab('consulta');
    await loadFiscalComprasRows();
  });
}

if (elements.fiscalComprasAddItem) {
  elements.fiscalComprasAddItem.addEventListener('click', () => {
    if (!canFiscalComprasCreate() && !canFiscalComprasEdit()) {
      showToast('Usuário sem permissão para editar itens.', true);
      return;
    }
    const items = readFiscalComprasItemsFromForm();
    items.push(createFiscalCompraItemSeed());
    renderFiscalComprasItems(items);
    recalcFiscalComprasTotalForm();
  });
}

if (elements.fiscalComprasCancel) {
  elements.fiscalComprasCancel.addEventListener('click', () => {
    resetFiscalComprasForm();
  });
}

async function saveFiscalCompraFromForm({ generatePdf = false } = {}) {
  try {
    const id = String(elements.fiscalComprasId?.value || '').trim();
    const payload = {
      username: getAuthUsername(),
      id,
      data_solicitacao: String(elements.fiscalComprasDataSolic?.value || '').trim(),
      setor: String(elements.fiscalComprasSetor?.value || '').trim(),
      destinado_a: String(elements.fiscalComprasDestinado?.value || '').trim(),
      observacoes: String(elements.fiscalComprasObservacoes?.value || '').trim(),
      status: String(elements.fiscalComprasStatus?.value || 'PENDENTE').trim(),
      itens: readFiscalComprasItemsFromForm()
    };
    if (elements.fiscalComprasError) {
      elements.fiscalComprasError.textContent = '';
    }
    if (elements.fiscalComprasSave) elements.fiscalComprasSave.disabled = true;
    if (elements.fiscalComprasSavePdf) elements.fiscalComprasSavePdf.disabled = true;

    const result = await window.api.fiscalComprasUpsert(payload);
    const rowId = result && result.row ? result.row.id : '';
    if (generatePdf && canFiscalComprasPdf() && rowId) {
      await window.api.fiscalComprasPdfRegen({ username: getAuthUsername(), id: rowId });
    }
    await loadFiscalComprasRows();
    showToast(generatePdf ? 'Solicitação salva e PDF gerado.' : 'Solicitação salva.');
    if (result && result.row) {
      fillFiscalComprasForm(result.row);
    }
    setFiscalComprasTab('consulta');
  } catch (error) {
    if (elements.fiscalComprasError) {
      elements.fiscalComprasError.textContent = error.message || 'Erro ao salvar solicitação.';
    } else {
      showToast(error.message || 'Erro ao salvar solicitação.', true);
    }
  } finally {
    applyFiscalComprasPermissionsUi();
  }
}

if (elements.fiscalComprasSave) {
  elements.fiscalComprasSave.addEventListener('click', async () => {
    await saveFiscalCompraFromForm({ generatePdf: false });
  });
}
if (elements.fiscalComprasSavePdf) {
  elements.fiscalComprasSavePdf.addEventListener('click', async () => {
    await saveFiscalCompraFromForm({ generatePdf: true });
  });
}

if (elements.fiscalComprasFilterClear) {
  elements.fiscalComprasFilterClear.addEventListener('click', async () => {
    if (elements.fiscalComprasFilterIni) elements.fiscalComprasFilterIni.value = '';
    if (elements.fiscalComprasFilterFim) elements.fiscalComprasFilterFim.value = '';
    if (elements.fiscalComprasFilterSolicitante) elements.fiscalComprasFilterSolicitante.value = '';
    if (elements.fiscalComprasFilterSetor) elements.fiscalComprasFilterSetor.value = '';
    if (elements.fiscalComprasFilterNumero) elements.fiscalComprasFilterNumero.value = '';
    if (elements.fiscalComprasFilterFornecedor) elements.fiscalComprasFilterFornecedor.value = '';
    if (elements.fiscalComprasFilterItem) elements.fiscalComprasFilterItem.value = '';
    if (elements.fiscalComprasFilterStatus) elements.fiscalComprasFilterStatus.value = '';
    state.fiscalComprasFilters = {};
    await loadFiscalComprasRows();
  });
}

if (elements.fiscalComprasFilterApply) {
  elements.fiscalComprasFilterApply.addEventListener('click', async () => {
    state.fiscalComprasFilters = {
      data_ini: String(elements.fiscalComprasFilterIni?.value || '').trim(),
      data_fim: String(elements.fiscalComprasFilterFim?.value || '').trim(),
      solicitante: String(elements.fiscalComprasFilterSolicitante?.value || '').trim(),
      setor: String(elements.fiscalComprasFilterSetor?.value || '').trim(),
      numero_requisicao: String(elements.fiscalComprasFilterNumero?.value || '').trim(),
      fornecedor: String(elements.fiscalComprasFilterFornecedor?.value || '').trim(),
      item_descricao: String(elements.fiscalComprasFilterItem?.value || '').trim(),
      status: String(elements.fiscalComprasFilterStatus?.value || '').trim()
    };
    await loadFiscalComprasRows();
  });
}

if (elements.userModalConfirm) {
  elements.userModalConfirm.addEventListener('click', async () => {
    try {
      const username = String(elements.userUsername?.value || '').trim();
      const password = String(elements.userPassword?.value || '').trim();
      const permissions = collectUserModalPermissions();
      if (!permissions.fiscal) {
        FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
          permissions[perm.key] = false;
        });
      } else {
        const hasComprasEnabled = FISCAL_COMPRAS_PERMISSION_DEFS.some((perm) => !!permissions[perm.key]);
        if (!hasComprasEnabled) {
          FISCAL_COMPRAS_PERMISSION_DEFS.forEach((perm) => {
            permissions[perm.key] = true;
          });
        }
      }

      if (!username || !password) {
        elements.userModalError.textContent = 'Informe usuário e senha.';
        return;
      }

      await window.api.createUser({
        username,
        password,
        permissions,
        adminPassword: state.adminSettingsPassword
      });

      closeUserModal();
      await loadUsers();
      renderUsersTable();
      applyPermissionsToTabs();
      showToast('Login criado.');
    } catch (error) {
      elements.userModalError.textContent = error.message || 'Erro ao criar login.';
    }
  });
}
if (elements.userModalCancel) {
  elements.userModalCancel.addEventListener('click', () => closeUserModal());
}
if (elements.userModalBackdrop) {
  elements.userModalBackdrop.addEventListener('click', () => closeUserModal());
}

if (elements.userPassConfirm) {
  elements.userPassConfirm.addEventListener('click', async () => {
    try {
      const username = String(state.userPassTarget || '').trim();
      const password = String(elements.userPassInput?.value || '').trim();
      if (!username) {
        return;
      }
      if (!password) {
        elements.userPassError.textContent = 'Informe a nova senha.';
        return;
      }

      await window.api.setUserPassword({ username, password, adminPassword: state.adminSettingsPassword });
      closeUserPassModal();
      showToast('Senha atualizada.');
    } catch (error) {
      elements.userPassError.textContent = error.message || 'Erro ao atualizar senha.';
    }
  });
}
if (elements.userPassCancel) {
  elements.userPassCancel.addEventListener('click', () => closeUserPassModal());
}
if (elements.userPassBackdrop) {
  elements.userPassBackdrop.addEventListener('click', () => closeUserPassModal());
}

if (elements.adminSettingsConfirm) {
  elements.adminSettingsConfirm.addEventListener('click', async () => {
    const password = String(elements.adminSettingsPassword?.value || '').trim();
    if (!password) {
      elements.adminSettingsError.textContent = 'Informe a senha administrativa.';
      return;
    }

    if (password !== SETTINGS_UNLOCK_PASSWORD) {
      elements.adminSettingsError.textContent = 'Senha administrativa incorreta.';
      elements.adminSettingsPassword?.focus?.();
      elements.adminSettingsPassword?.select?.();
      return;
    }

    state.adminSettingsPassword = password;
    closeAdminSettingsModal();

    const action = state.adminPendingAction;
    state.adminPendingAction = null;
    if (typeof action === 'function') {
      await action();
    }
  });
}
if (elements.adminSettingsCancel) {
  elements.adminSettingsCancel.addEventListener('click', () => {
    state.adminPendingAction = null;
    closeAdminSettingsModal();
  });
}
if (elements.adminSettingsBackdrop) {
  elements.adminSettingsBackdrop.addEventListener('click', () => {
    state.adminPendingAction = null;
    closeAdminSettingsModal();
  });
}
if (elements.adminSettingsPassword) {
  elements.adminSettingsPassword.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      elements.adminSettingsConfirm?.click?.();
    } else if (event.key === 'Escape') {
      state.adminPendingAction = null;
      closeAdminSettingsModal();
    }
  });
}

async function openRhHeTool() {
  try {
    ensureRhMonthValue();
    await refreshRhAll();
    setRhView('he');
    setRhHeSubView('dashboard');
  } catch (error) {
    showToast(error.message || 'Erro ao abrir RH.', true);
  }
}

function setRhSolicHeTab(next) {
  const tab = String(next || '').trim() === 'consulta' ? 'consulta' : 'form';
  state.rhSolicHeTab = tab;
  if (elements.rhSolicHeViewForm) {
    elements.rhSolicHeViewForm.hidden = tab !== 'form';
  }
  if (elements.rhSolicHeViewConsulta) {
    elements.rhSolicHeViewConsulta.hidden = tab !== 'consulta';
  }
  elements.rhSolicHeTabForm?.classList.toggle('active', tab === 'form');
  elements.rhSolicHeTabConsulta?.classList.toggle('active', tab === 'consulta');
}

function computeSolicHeTotalMinutesFromTimes(start, end) {
  const sText = String(start || '').trim();
  const eText = String(end || '').trim();
  const matchStart = sText.match(/^(\d{1,2}):(\d{2})$/);
  const matchEnd = eText.match(/^(\d{1,2}):(\d{2})$/);
  if (!matchStart || !matchEnd) {
    return null;
  }
  const s = (Number(matchStart[1]) * 60) + Number(matchStart[2]);
  const e = (Number(matchEnd[1]) * 60) + Number(matchEnd[2]);
  if (!Number.isFinite(s) || !Number.isFinite(e) || e <= s) {
    return null;
  }
  return e - s;
}

function formatSolicHeTotalHoursTextFromTimes(start, end) {
  const diffMin = computeSolicHeTotalMinutesFromTimes(start, end);
  if (diffMin == null) {
    return '';
  }
  const h = Math.floor(diffMin / 60);
  const m = diffMin % 60;
  return `${h}:${String(m).padStart(2, '0')}`;
}

function computeSolicHeTotalHoursText() {
  return formatSolicHeTotalHoursTextFromTimes(
    elements.rhSolicHeInicio?.value,
    elements.rhSolicHeFim?.value
  );
}

function getSelectedSolicHeColab() {
  const id = String(elements.rhSolicHeColab?.value || '').trim();
  if (!id) {
    return null;
  }
  return (state.rhColaboradores || []).find((c) => String(c.id) === id) || null;
}

function clearRhSolicHeForm({ keepSolicitante = true } = {}) {
  if (elements.rhSolicHeError) {
    elements.rhSolicHeError.textContent = '';
  }
  elements.rhSolicHeId && (elements.rhSolicHeId.value = '');
  elements.rhSolicHeNumero && (elements.rhSolicHeNumero.value = '');
  if (elements.rhSolicHeDataHe) {
    elements.rhSolicHeDataHe.value = getTodayYmdLocal();
  }
  if (elements.rhSolicHeFinalidade) {
    elements.rhSolicHeFinalidade.value = '';
  }
  if (elements.rhSolicHeInicio) {
    elements.rhSolicHeInicio.value = '';
  }
  if (elements.rhSolicHeFim) {
    elements.rhSolicHeFim.value = '';
  }
  if (elements.rhSolicHeTotal) {
    elements.rhSolicHeTotal.value = '';
  }
  if (elements.rhSolicHeSetor) {
    elements.rhSolicHeSetor.value = '';
  }
  if (elements.rhSolicHeSolicitante) {
    elements.rhSolicHeSolicitante.value = keepSolicitante ? getAuthUsername() : '';
  }
  if (elements.rhSolicHeDataSolic) {
    elements.rhSolicHeDataSolic.value = formatYmdToBr(getTodayYmdLocal());
  }
}

function fillRhSolicHeForm(row) {
  clearRhSolicHeForm();
  if (elements.rhSolicHeId) elements.rhSolicHeId.value = String(row.id || '').trim();
  if (elements.rhSolicHeNumero) elements.rhSolicHeNumero.value = String(row.numero_documento || '').trim();
  if (elements.rhSolicHeDataHe) elements.rhSolicHeDataHe.value = String(row.data_he || '').trim();
  if (elements.rhSolicHeFinalidade) elements.rhSolicHeFinalidade.value = String(row.finalidade || '').trim();
  if (elements.rhSolicHeInicio) elements.rhSolicHeInicio.value = String(row.hora_inicio || '').trim();
  if (elements.rhSolicHeFim) elements.rhSolicHeFim.value = String(row.hora_fim || '').trim();
  if (elements.rhSolicHeTotal) elements.rhSolicHeTotal.value = computeSolicHeTotalHoursText();
  if (elements.rhSolicHeSolicitante) elements.rhSolicHeSolicitante.value = String(row.solicitante || '').trim();
  if (elements.rhSolicHeDataSolic) elements.rhSolicHeDataSolic.value = formatYmdToBr(String(row.data_solicitacao || '').trim());

  if (elements.rhSolicHeColab) {
    elements.rhSolicHeColab.value = String(row.colaborador_id || '').trim();
  }
  const colab = getSelectedSolicHeColab();
  if (elements.rhSolicHeSetor) {
    elements.rhSolicHeSetor.value = colab ? (getSetorLabelById(colab.setor_id) || String(row.setor || '').trim()) : String(row.setor || '').trim();
  }
}

async function refreshRhSolicHeList() {
  const username = getAuthUsername();
  const filters = state.rhSolicHeFilters || {};
  const result = await window.api.rhSolicHeList({ username, filters });
  state.rhSolicHeRows = (result && result.rows) ? result.rows : [];
  renderRhSolicHeTable();
}

function renderRhSolicHeTable() {
  if (!elements.rhSolicHeTableHead || !elements.rhSolicHeTableBody) {
    return;
  }
  elements.rhSolicHeTableHead.innerHTML = '<tr><th>Nº</th><th>Data H.E</th><th>Colaborador</th><th>Setor</th><th>Total</th><th>Solicitante</th><th>Data Solic.</th><th>Ações</th></tr>';
  elements.rhSolicHeTableBody.innerHTML = '';

  const rows = Array.isArray(state.rhSolicHeRows) ? state.rhSolicHeRows : [];
  if (elements.rhSolicHeConsultaEmpty) {
    elements.rhSolicHeConsultaEmpty.hidden = rows.length > 0;
  }

  rows.forEach((r) => {
    const tr = document.createElement('tr');
    const tdDoc = document.createElement('td');
    tdDoc.textContent = String(r.numero_documento || '').trim();
    const tdData = document.createElement('td');
    tdData.textContent = formatYmdToBr(String(r.data_he || '').trim());
    const tdColab = document.createElement('td');
    tdColab.textContent = String(r.nome_colaborador || '').trim();
    const tdSetor = document.createElement('td');
    tdSetor.textContent = String(r.setor || '').trim();
    const tdTot = document.createElement('td');
    tdTot.textContent = formatSolicHeTotalHoursTextFromTimes(r.hora_inicio, r.hora_fim) || formatHours(Number(r.total_horas || 0));
    const tdSolic = document.createElement('td');
    tdSolic.textContent = String(r.solicitante || '').trim();
    const tdDataSol = document.createElement('td');
    tdDataSol.textContent = formatYmdToBr(String(r.data_solicitacao || '').trim());

    const tdAcoes = document.createElement('td');
    const viewBtn = document.createElement('button');
    viewBtn.className = 'btn secondary';
    viewBtn.type = 'button';
    viewBtn.textContent = 'Visualizar';
    viewBtn.addEventListener('click', async () => {
      try {
        await window.api.rhSolicHePdfOpen({ username: getAuthUsername(), id: r.id });
      } catch (error) {
        showToast(error.message || 'Erro ao abrir PDF.', true);
      }
    });

    const editBtn = document.createElement('button');
    editBtn.className = 'btn secondary';
    editBtn.type = 'button';
    editBtn.textContent = 'Editar';
    editBtn.style.marginLeft = '8px';
    editBtn.addEventListener('click', () => {
      fillRhSolicHeForm(r);
      setRhSolicHeTab('form');
    });

    const pdfBtn = document.createElement('button');
    pdfBtn.className = 'btn secondary';
    pdfBtn.type = 'button';
    pdfBtn.textContent = 'PDF';
    pdfBtn.style.marginLeft = '8px';
    pdfBtn.addEventListener('click', async () => {
      try {
        await window.api.rhSolicHePdfRegen({ username: getAuthUsername(), id: r.id });
        showToast('PDF gerado novamente.');
      } catch (error) {
        showToast(error.message || 'Erro ao gerar PDF.', true);
      }
    });

    const delBtn = document.createElement('button');
    delBtn.className = 'btn danger';
    delBtn.type = 'button';
    delBtn.textContent = 'Excluir';
    delBtn.style.marginLeft = '8px';
    delBtn.addEventListener('click', () => confirmRhAction('Excluir esta solicitação e o PDF?', async () => {
      await window.api.rhSolicHeDelete({ username: getAuthUsername(), id: r.id });
      await refreshRhSolicHeList();
    }));

    tdAcoes.appendChild(viewBtn);
    tdAcoes.appendChild(editBtn);
    tdAcoes.appendChild(delBtn);
    tdAcoes.appendChild(pdfBtn);

    tr.appendChild(tdDoc);
    tr.appendChild(tdData);
    tr.appendChild(tdColab);
    tr.appendChild(tdSetor);
    tr.appendChild(tdTot);
    tr.appendChild(tdSolic);
    tr.appendChild(tdDataSol);
    tr.appendChild(tdAcoes);
    elements.rhSolicHeTableBody.appendChild(tr);
  });
}

async function openRhSolicHeTool() {
  try {
    ensureRhSolicHeApiAvailable();
    await loadRhContext();
    await loadRhBaseLists();

    if (elements.rhSolicHeColab) {
      const active = (state.rhColaboradores || []).filter((c) => String(c.status || '').trim() !== 'inativo');
      setRhSelectOptions(
        elements.rhSolicHeColab,
        active.map((c) => ({ value: c.id, label: c.nome })),
        { includeAll: true, allLabel: 'Selecione' }
      );
    }
    if (elements.rhSolicHeSolicitante) {
      elements.rhSolicHeSolicitante.value = getAuthUsername();
    }
    if (elements.rhSolicHeDataSolic) {
      elements.rhSolicHeDataSolic.value = formatYmdToBr(getTodayYmdLocal());
    }

    // Filtros
    setRhSelectOptions(
      elements.rhSolicHeFilterColab,
      (state.rhColaboradores || []).map((c) => ({ value: c.id, label: c.nome })),
      { includeAll: true, allLabel: 'Todos' }
    );
    setRhSelectOptions(
      elements.rhSolicHeFilterSetor,
      (state.rhSetores || []).map((s) => ({ value: s.nome_setor || s.nome || s.id, label: s.nome_setor || s.nome || s.id })),
      { includeAll: true, allLabel: 'Todos' }
    );

    clearRhSolicHeForm();
    setRhSolicHeTab('form');
    state.rhSolicHeFilters = {};
    await refreshRhSolicHeList();

    setRhView('solic_he');
  } catch (error) {
    showToast(error.message || 'Erro ao abrir solicitações.', true);
  }
}

async function handleRhSubtab(next) {
  setRhHeSubView(next);
  try {
    if (next === 'config') {
      await renderRhConfig();
    } else if (next === 'audit') {
      await renderRhAudit();
    }
  } catch (error) {
    showToast(error.message || 'Erro ao carregar visão RH.', true);
  }
}

async function handleRhExportExcel() {
  try {
    const username = getAuthUsername();
    const filters = getRhUiFilters();
    const result = await window.api.rhExportExcel({
      username,
      ...filters,
      suggestedName: `rh-he-${filters.month}.xlsx`
    });
    if (result && result.canceled) {
      return;
    }
    showToast('Excel exportado com sucesso.');
  } catch (error) {
    showToast(error.message || 'Erro ao exportar Excel.', true);
  }
}

async function handleRhExportPdf() {
  try {
    const username = getAuthUsername();
    const filters = getRhUiFilters();
    setRhHeSubView('dashboard');
    elements.body.classList.add('pdf-rh-export');
    const result = await window.api.rhExportPdf({
      suggestedName: `rh-he-${filters.month}.pdf`,
      dialogTitle: 'Salvar exportação RH em PDF',
      exportWidth: 1122,
      exportHeight: 794
    });
    if (!result || result.canceled) {
      return;
    }
    showToast('PDF exportado com sucesso.');
  } catch (error) {
    showToast(error.message || 'Erro ao exportar PDF.', true);
  } finally {
    elements.body.classList.remove('pdf-rh-export');
  }
}

if (elements.rhOpenHe) {
  elements.rhOpenHe.addEventListener('click', openRhHeTool);
}
if (elements.rhOpenSolicHe) {
  elements.rhOpenSolicHe.addEventListener('click', openRhSolicHeTool);
}
if (elements.rhHeBackMenu) {
  elements.rhHeBackMenu.addEventListener('click', () => setRhView('menu'));
}
if (elements.rhSolicHeBackMenu) {
  elements.rhSolicHeBackMenu.addEventListener('click', () => setRhView('menu'));
}
if (elements.rhHeRefresh) {
  elements.rhHeRefresh.addEventListener('click', async () => {
    try {
      await refreshRhAll();
      showToast('RH atualizado.');
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar RH.', true);
    }
  });
}
if (elements.rhSolicHeRefresh) {
  elements.rhSolicHeRefresh.addEventListener('click', async () => {
    try {
      await refreshRhSolicHeList();
      showToast('Solicitações atualizadas.');
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar solicitações.', true);
    }
  });
}
if (elements.rhSolicHeNew) {
  elements.rhSolicHeNew.addEventListener('click', () => {
    clearRhSolicHeForm({ keepSolicitante: true });
    setRhSolicHeTab('form');
  });
}
if (elements.rhSolicHeTabForm) {
  elements.rhSolicHeTabForm.addEventListener('click', () => setRhSolicHeTab('form'));
}
if (elements.rhSolicHeTabConsulta) {
  elements.rhSolicHeTabConsulta.addEventListener('click', () => setRhSolicHeTab('consulta'));
}
if (elements.rhSolicHeColab) {
  elements.rhSolicHeColab.addEventListener('change', () => {
    const colab = getSelectedSolicHeColab();
    if (elements.rhSolicHeSetor) {
      elements.rhSolicHeSetor.value = colab ? (getSetorLabelById(colab.setor_id) || '') : '';
    }
  });
}
['rhSolicHeInicio', 'rhSolicHeFim'].forEach((key) => {
  const el = elements[key];
  if (el) {
    el.addEventListener('change', () => {
      if (elements.rhSolicHeTotal) {
        elements.rhSolicHeTotal.value = computeSolicHeTotalHoursText();
      }
    });
  }
});
if (elements.rhSolicHeCancel) {
  elements.rhSolicHeCancel.addEventListener('click', () => {
    clearRhSolicHeForm({ keepSolicitante: true });
  });
}
if (elements.rhSolicHeSave) {
  elements.rhSolicHeSave.addEventListener('click', async () => {
    try {
      const username = getAuthUsername();
      const id = String(elements.rhSolicHeId?.value || '').trim();
      const data_he = String(elements.rhSolicHeDataHe?.value || '').trim();
      const colaborador_id = String(elements.rhSolicHeColab?.value || '').trim();
      const colab = getSelectedSolicHeColab();
      const nome_colaborador = colab ? String(colab.nome || '').trim() : '';
      const setor = colab ? (getSetorLabelById(colab.setor_id) || '') : String(elements.rhSolicHeSetor?.value || '').trim();
      const finalidade = String(elements.rhSolicHeFinalidade?.value || '').trim();
      const hora_inicio = String(elements.rhSolicHeInicio?.value || '').trim();
      const hora_fim = String(elements.rhSolicHeFim?.value || '').trim();

      if (!id) {
        if (elements.rhSolicHeSolicitante) {
          elements.rhSolicHeSolicitante.value = getAuthUsername();
        }
        if (elements.rhSolicHeDataSolic) {
          elements.rhSolicHeDataSolic.value = formatYmdToBr(getTodayYmdLocal());
        }
      }

      if (!data_he) {
        throw new Error('Informe a data da Hora Extra.');
      }
      if (!colaborador_id) {
        throw new Error('Selecione o colaborador.');
      }
      if (!setor) {
        throw new Error('Setor não identificado.');
      }
      if (!finalidade) {
        throw new Error('Informe a finalidade.');
      }
      if (!hora_inicio || !hora_fim) {
        throw new Error('Informe o horário de início e término.');
      }

      if (elements.rhSolicHeError) {
        elements.rhSolicHeError.textContent = '';
      }
      elements.rhSolicHeSave.disabled = true;

      const result = await window.api.rhSolicHeUpsert({
        username,
        id,
        data_he,
        colaborador_id,
        nome_colaborador,
        setor,
        finalidade,
        hora_inicio,
        hora_fim
      });

      const row = result && result.row ? result.row : null;
      if (row) {
        fillRhSolicHeForm(row);
      }
      showToast('Solicitação salva e PDF gerado.');
      await refreshRhSolicHeList();
      setRhSolicHeTab('consulta');
    } catch (error) {
      if (elements.rhSolicHeError) {
        elements.rhSolicHeError.textContent = error.message || 'Erro ao salvar solicitação.';
      } else {
        showToast(error.message || 'Erro ao salvar solicitação.', true);
      }
    } finally {
      if (elements.rhSolicHeSave) {
        elements.rhSolicHeSave.disabled = false;
      }
    }
  });
}

if (elements.rhSolicHeFilterClear) {
  elements.rhSolicHeFilterClear.addEventListener('click', async () => {
    if (elements.rhSolicHeFilterColab) elements.rhSolicHeFilterColab.value = '';
    if (elements.rhSolicHeFilterSetor) elements.rhSolicHeFilterSetor.value = '';
    if (elements.rhSolicHeFilterIni) elements.rhSolicHeFilterIni.value = '';
    if (elements.rhSolicHeFilterFim) elements.rhSolicHeFilterFim.value = '';
    if (elements.rhSolicHeFilterSolicitante) elements.rhSolicHeFilterSolicitante.value = '';
    if (elements.rhSolicHeFilterDoc) elements.rhSolicHeFilterDoc.value = '';
    state.rhSolicHeFilters = {};
    await refreshRhSolicHeList();
  });
}
if (elements.rhSolicHeFilterApply) {
  elements.rhSolicHeFilterApply.addEventListener('click', async () => {
    state.rhSolicHeFilters = {
      colaborador_id: String(elements.rhSolicHeFilterColab?.value || '').trim(),
      setor: String(elements.rhSolicHeFilterSetor?.value || '').trim(),
      data_ini: String(elements.rhSolicHeFilterIni?.value || '').trim(),
      data_fim: String(elements.rhSolicHeFilterFim?.value || '').trim(),
      solicitante: String(elements.rhSolicHeFilterSolicitante?.value || '').trim(),
      numero_documento: String(elements.rhSolicHeFilterDoc?.value || '').trim()
    };
    await refreshRhSolicHeList();
  });
}
if (elements.rhHeAdd) {
  elements.rhHeAdd.addEventListener('click', () => {
    if (!canEditRh()) {
      showToast('Somente RH pode lançar/editar.', true);
      return;
    }
    openRhHeModal(null);
  });
}
if (elements.rhHeImport) {
  elements.rhHeImport.addEventListener('click', () => {
    if (!canEditRh()) {
      showToast('Somente RH (edição) pode importar.', true);
      return;
    }
    openRhImportModal();
  });
}
if (elements.rhHeExportExcel) {
  elements.rhHeExportExcel.addEventListener('click', handleRhExportExcel);
}
if (elements.rhHeExportPdf) {
  elements.rhHeExportPdf.addEventListener('click', handleRhExportPdf);
}
if (elements.rhFilterMonth) {
  elements.rhFilterMonth.addEventListener('change', async () => {
    try {
      await refreshRhAll();
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar filtros.', true);
    }
  });
}
if (elements.rhFilterSector) {
  elements.rhFilterSector.addEventListener('change', async () => {
    try {
      await refreshRhAll();
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar filtros.', true);
    }
  });
}
if (elements.rhFilterCollab) {
  elements.rhFilterCollab.addEventListener('change', async () => {
    try {
      await refreshRhAll();
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar filtros.', true);
    }
  });
}
if (elements.rhFilterType) {
  elements.rhFilterType.addEventListener('change', async () => {
    try {
      await refreshRhAll();
    } catch (error) {
      showToast(error.message || 'Erro ao atualizar filtros.', true);
    }
  });
}

[
  { el: elements.rhSubtabDashboard, view: 'dashboard' },
  { el: elements.rhSubtabAtrasos, view: 'atrasos' },
  { el: elements.rhSubtabRegistros, view: 'registros' },
  { el: elements.rhSubtabColaboradores, view: 'colaboradores' },
  { el: elements.rhSubtabSetores, view: 'setores' },
  { el: elements.rhSubtabConfig, view: 'config' },
  { el: elements.rhSubtabAudit, view: 'audit' }
].forEach(({ el, view }) => {
  if (el) {
    el.addEventListener('click', () => handleRhSubtab(view));
  }
});

if (elements.rhSectorCancel) {
  elements.rhSectorCancel.addEventListener('click', closeRhSectorModal);
}
if (elements.rhSectorBackdrop) {
  elements.rhSectorBackdrop.addEventListener('click', closeRhSectorModal);
}
if (elements.rhSectorConfirm) {
  elements.rhSectorConfirm.addEventListener('click', async () => {
    try {
      const username = getAuthUsername();
      const id = String(elements.rhSectorId?.value || '').trim();
      const nome_setor = String(elements.rhSectorNome?.value || '').trim();
      const gestor_responsavel = String(elements.rhSectorGestor?.value || '').trim();
      const centro_custo = String(elements.rhSectorCc?.value || '').trim();
      if (!nome_setor) {
        throw new Error('Informe o nome do setor.');
      }
      await window.api.rhUpsertSetor({ username, id, nome_setor, gestor_responsavel, centro_custo });
      closeRhSectorModal();
      await refreshRhAll();
      setRhHeSubView('setores');
      showToast('Setor salvo.');
    } catch (error) {
      if (elements.rhSectorError) {
        elements.rhSectorError.textContent = error.message || 'Erro ao salvar setor.';
      }
    }
  });
}

if (elements.rhCollabCancel) {
  elements.rhCollabCancel.addEventListener('click', closeRhCollabModal);
}
if (elements.rhCollabBackdrop) {
  elements.rhCollabBackdrop.addEventListener('click', closeRhCollabModal);
}
if (elements.rhCollabConfirm) {
  elements.rhCollabConfirm.addEventListener('click', async () => {
    try {
      const username = getAuthUsername();
      const id = String(elements.rhCollabId?.value || '').trim();
      const nome = String(elements.rhCollabNome?.value || '').trim();
      const matricula = String(elements.rhCollabMatricula?.value || '').trim();
      const setor_id = String(elements.rhCollabSetor?.value || '').trim();
      const cargo = String(elements.rhCollabCargo?.value || '').trim();
      const status = String(elements.rhCollabStatus?.value || 'ativo').trim();
      const data_admissao = String(elements.rhCollabAdmissao?.value || '').trim();
      const limite_mensal = String(elements.rhCollabLimite?.value || '').trim();
      if (!nome) {
        throw new Error('Informe o nome do colaborador.');
      }
      await window.api.rhUpsertColaborador({ username, id, nome, matricula, setor_id, cargo, status, data_admissao, limite_mensal });
      closeRhCollabModal();
      await refreshRhAll();
      setRhHeSubView('colaboradores');
      showToast('Colaborador salvo.');
    } catch (error) {
      if (elements.rhCollabError) {
        elements.rhCollabError.textContent = error.message || 'Erro ao salvar colaborador.';
      }
    }
  });
}

if (elements.rhHeCancel) {
  elements.rhHeCancel.addEventListener('click', closeRhHeModal);
}
if (elements.rhHeBackdrop) {
  elements.rhHeBackdrop.addEventListener('click', closeRhHeModal);
}
if (elements.rhHeKind) {
  elements.rhHeKind.addEventListener('change', () => {
    setRhHeModalKind(String(elements.rhHeKind?.value || 'he').trim(), { lock: false });
  });
}
if (elements.rhHeConfirm) {
  elements.rhHeConfirm.addEventListener('click', async () => {
    try {
      const username = getAuthUsername();
      const id = String(elements.rhHeId?.value || '').trim();
      const kind = String(elements.rhHeKind?.value || 'he').trim() === 'atraso' ? 'atraso' : 'he';
      const colaborador_id = String(elements.rhHeColaborador?.value || '').trim();
      const data = String(elements.rhHeData?.value || '').trim();
      const quantidade_horas = parseDurationTextToHours(elements.rhHeQtd?.value);
      const tipo_hora = String(elements.rhHeTipo?.value || '').trim();
      const observacao = String(elements.rhHeObs?.value || '').trim();
      const justificativa = String(elements.rhHeJust?.value || '').trim();
      if (!colaborador_id) {
        throw new Error('Selecione um colaborador.');
      }
      if (!data) {
        throw new Error('Informe a data.');
      }
      if (!Number.isFinite(quantidade_horas) || quantidade_horas <= 0) {
        throw new Error('Informe a duração (ex: 1:30).');
      }
      if (kind === 'atraso') {
        if (!window.api || typeof window.api.rhUpsertAtraso !== 'function') {
          throw new Error('Função de atrasos não disponível (atualize o app).');
        }
        await window.api.rhUpsertAtraso({
          username,
          id,
          colaborador_id,
          data,
          quantidade_horas,
          observacao
        });
      } else {
        const limiteJust = Number(state.rhContext?.config?.justificativa_acima_horas ?? 2);
        if (Number.isFinite(limiteJust) && quantidade_horas > limiteJust && !justificativa) {
          throw new Error(`Justificativa obrigatória acima de ${limiteJust}h.`);
        }
        await window.api.rhUpsertHoraExtra({
          username,
          id,
          colaborador_id,
          data,
          quantidade_horas,
          tipo_hora,
          observacao,
          justificativa
        });
      }
      closeRhHeModal();
      await refreshRhAll();
      setRhHeSubView(kind === 'atraso' ? 'atrasos' : 'registros');
      showToast(kind === 'atraso' ? 'Atraso salvo.' : 'Lançamento salvo.');
    } catch (error) {
      if (elements.rhHeError) {
        elements.rhHeError.textContent = error.message || 'Erro ao salvar lançamento.';
      }
    }
  });
}

if (elements.rhConfirmCancel) {
  elements.rhConfirmCancel.addEventListener('click', () => {
    state.rhPendingConfirm = null;
    closeRhModal(elements.rhConfirmModal);
  });
}
if (elements.rhConfirmBackdrop) {
  elements.rhConfirmBackdrop.addEventListener('click', () => {
    state.rhPendingConfirm = null;
    closeRhModal(elements.rhConfirmModal);
  });
}
if (elements.rhConfirmOk) {
  elements.rhConfirmOk.addEventListener('click', async () => {
    const action = state.rhPendingConfirm;
    state.rhPendingConfirm = null;
    closeRhModal(elements.rhConfirmModal);
    if (typeof action === 'function') {
      try {
        await action();
      } catch (error) {
        showToast(error.message || 'Erro ao executar operação.', true);
      }
    }
  });
}

if (elements.rhImportCancel) {
  elements.rhImportCancel.addEventListener('click', closeRhImportModal);
}
if (elements.rhImportBackdrop) {
  elements.rhImportBackdrop.addEventListener('click', closeRhImportModal);
}
if (elements.rhImportConfirm) {
  elements.rhImportConfirm.addEventListener('click', handleRhImportConfirm);
}
if (elements.rhImportFilePick || elements.rhImportFilePath || elements.rhImportFilePicker) {
  const pickWorkbook = async () => {
    try {
      if (window.api && typeof window.api.rhPickImportWorkbook === 'function') {
        const picked = await window.api.rhPickImportWorkbook();
        if (picked) {
          setRhImportWorkbookPath(picked);
        }
        return;
      }
      elements.rhImportFile?.click?.();
    } catch (error) {
      showToast(error.message || 'Erro ao selecionar planilha.', true);
    }
  };

  if (elements.rhImportFilePick) {
    elements.rhImportFilePick.addEventListener('click', pickWorkbook);
  }
  if (elements.rhImportFilePath) {
    elements.rhImportFilePath.addEventListener('click', pickWorkbook);
  }
  if (elements.rhImportFilePicker) {
    elements.rhImportFilePicker.addEventListener('dragover', (event) => {
      event.preventDefault();
    });
    elements.rhImportFilePicker.addEventListener('drop', (event) => {
      event.preventDefault();
      const dropped = event.dataTransfer?.files?.[0]?.path ? String(event.dataTransfer.files[0].path) : '';
      if (dropped) {
        setRhImportWorkbookPath(dropped);
      }
    });
  }
}
if (elements.rhImportFile) {
  elements.rhImportFile.addEventListener('change', () => {
    const filePath = elements.rhImportFile?.files?.[0]?.path ? String(elements.rhImportFile.files[0].path) : '';
    if (filePath) {
      setRhImportWorkbookPath(filePath);
    }
  });
}

init();





