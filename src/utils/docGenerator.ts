/**
 * Gerador de Documento .DOC com Especificação de Requisitos e Arquitetura
 * do Sistema Nexus Política / Plataforma Eleitoral 2026.
 */

export function downloadRequirementsDoc() {
  const documentContent = `
<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <meta charset="utf-8">
  <title>Especificação de Requisitos e Arquitetura - Nexus Política 2026</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    body {
      font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.5;
      color: #1a1a1a;
      margin: 40pt;
    }
    h1 {
      font-size: 20pt;
      color: #1e3a8a;
      border-bottom: 2px solid #1e3a8a;
      padding-bottom: 6px;
      margin-top: 24pt;
      margin-bottom: 12pt;
      text-transform: uppercase;
    }
    h2 {
      font-size: 14pt;
      color: #2563eb;
      margin-top: 18pt;
      margin-bottom: 8pt;
      border-bottom: 1px solid #e5e7eb;
      padding-bottom: 4px;
    }
    h3 {
      font-size: 12pt;
      color: #1f2937;
      margin-top: 12pt;
      margin-bottom: 6pt;
    }
    p, li {
      font-size: 11pt;
      text-align: justify;
    }
    ul, ol {
      margin-top: 4pt;
      margin-bottom: 8pt;
      padding-left: 20pt;
    }
    li {
      margin-bottom: 4pt;
    }
    .header-box {
      background-color: #f8fafc;
      border: 1px solid #cbd5e1;
      padding: 16pt;
      border-radius: 4px;
      margin-bottom: 20pt;
    }
    .header-box h1 {
      margin-top: 0;
      color: #0f172a;
      border-bottom: none;
      padding-bottom: 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10pt;
      margin-bottom: 15pt;
    }
    th, td {
      border: 1px solid #cbd5e1;
      padding: 8pt 10pt;
      text-align: left;
      font-size: 10pt;
    }
    th {
      background-color: #f1f5f9;
      color: #0f172a;
      font-weight: bold;
    }
    tr:nth-child(even) {
      background-color: #f8fafc;
    }
    .highlight-box {
      background-color: #eff6ff;
      border-left: 4px solid #2563eb;
      padding: 10pt 12pt;
      margin: 12pt 0;
    }
    .footer-text {
      margin-top: 30pt;
      font-size: 9pt;
      color: #64748b;
      text-align: center;
      border-top: 1px solid #e2e8f0;
      padding-top: 10pt;
    }
  </style>
</head>
<body>

  <div class="header-box">
    <h1>DOCUMENTO DE ESPECIFICAÇÃO DE REQUISITOS E ARQUITETURA</h1>
    <p><strong>SISTEMA:</strong> NEXUS POLÍTICA 2026 — Plataforma de Gestão Tática e Inteligência Eleitoral</p>
    <p><strong>VERSÃO:</strong> 2.5 (Enterprise Production)</p>
    <p><strong>DATA DE EMISSÃO:</strong> Agosto de 2026</p>
    <p><strong>STATUS:</strong> Homologado &amp; Em Operação em Campo</p>
  </div>

  <h2>1. VISÃO GERAL DO SISTEMA</h2>
  <p>
    O <strong>Nexus Política 2026</strong> é uma plataforma full-stack de inteligência e gestão eleitoral projetada para coordenar campanhas políticas, mapear eleitores fidelizados, otimizar a distribuição de recursos em campo (combustível, material de campanha) e integrar dados do Tribunal Superior Eleitoral (TRE) com suporte a operação totalmente offline (PWA).
  </p>
  <p>
    A solução permite o acompanhamento em tempo real das atividades de campo, gestão de equipes táticas hierarquizadas (Coordenadores Gerais, Coordenadores Regionais, Líderes de Campo/Cabos Eleitorais), bem como o cadastramento público autônomo por eleitores via links exclusivos de indicação.
  </p>

  <h2>2. ARQUITETURA DA SOLUÇÃO</h2>
  
  <div class="highlight-box">
    <p><strong>Modelo Arquitetural:</strong> Single Page Application (SPA) reativa e escalável com arquitetura serverless em nuvem, banco de dados orientado a documentos em tempo real e sincronismo bidirecional PWA/IndexedDB para resiliência offline.</p>
  </div>

  <h3>2.1 Camadas do Sistema</h3>
  <table>
    <thead>
      <tr>
        <th>Camada</th>
        <th>Tecnologia / Framework</th>
        <th>Função Arquitetural</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>Frontend User Interface</strong></td>
        <td>React 18, TypeScript, Vite, Tailwind CSS</td>
        <td>Interface reativa, responsiva e otimizada para smartphones e desktops. Componentização modular baseada em atomic design.</td>
      </tr>
      <tr>
        <td><strong>Animações &amp; Layout</strong></td>
        <td>Motion (framer-motion), Lucide React</td>
        <td>Transições fluidas de tela, microinterações táticas e iconografia padronizada.</td>
      </tr>
      <tr>
        <td><strong>Backend &amp; Database</strong></td>
        <td>Firebase Cloud Firestore</td>
        <td>Banco de dados NoSQL distribuído com listeners reativos em tempo real ('onSnapshot'), regras de acesso por perfil e alta disponibilidade.</td>
      </tr>
      <tr>
        <td><strong>Autenticação &amp; Segurança</strong></td>
        <td>Firebase Authentication (Custom Claims / RBAC)</td>
        <td>Gestão de credenciais, sessão persistente, troca obrigatória de senha e níveis de permissão por papel (Role-Based Access Control).</td>
      </tr>
      <tr>
        <td><strong>Persistência Offline (PWA)</strong></td>
        <td>IndexedDB / LocalStorage Encapsulado ('eleitoralStorage')</td>
        <td>Fila de sincronização criptografada ('offline_voter_queue'), garantindo cadastro de eleitores e suporte mesmo sem sinal de internet.</td>
      </tr>
      <tr>
        <td><strong>Inteligência Artificial</strong></td>
        <td>Google Gemini API ('@google/genai')</td>
        <td>Geração automatizada de análises táticas, pareceres de viabilidade eleitoral, resumos de discursos e inteligência estratégica.</td>
      </tr>
      <tr>
        <td><strong>Visualização Geográfica</strong></td>
        <td>Componentes Customizados Roraima/TRE (D3.js / SVG)</td>
        <td>Mapeamento tático por município, zonas eleitorais e seções votantes do TRE.</td>
      </tr>
      <tr>
        <td><strong>Relatórios &amp; Exportação</strong></td>
        <td>SheetJS ('xlsx'), jsPDF, html2canvas</td>
        <td>Geração de planilhas de contingência em Excel, cartões de eleitor em PDF e modelos de importação em lote.</td>
      </tr>
    </tbody>
  </table>

  <h2>3. REQUISITOS FUNCIONAIS (RF)</h2>
  <table>
    <thead>
      <tr>
        <th>Código</th>
        <th>Nome do Requisito</th>
        <th>Descrição Funcional</th>
        <th>Prioridade</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>RF01</strong></td>
        <td>Gestão de Acesso e Perfis (RBAC)</td>
        <td>Controle estrito de privilégios para Administrador Geral, Coordenador Regional, Líder/Cabo Eleitoral e Eleitor. Suporte a bloqueio por plano e alteração obrigatória de senha.</td>
        <td>Crítica</td>
      </tr>
      <tr>
        <td><strong>RF02</strong></td>
        <td>Mapeamento de Eleitores em Campo</td>
        <td>Cadastro individual de eleitores com captura de nome, WhatsApp, CPF, título de eleitor, zona, seção, bairro, município e tagging de afinidade.</td>
        <td>Crítica</td>
      </tr>
      <tr>
        <td><strong>RF03</strong></td>
        <td>Cadastro em Lote via Planilha Excel</td>
        <td>Leitura, validação sintática e importação em massa de listas de eleitores através de arquivos '.xlsx' pré-formatados com tratamento de erros.</td>
        <td>Alta</td>
      </tr>
      <tr>
        <td><strong>RF04</strong></td>
        <td>Prevenção de Duplicidade Eleitoral</td>
        <td>Verificação automática no Firestore por número de telefone e CPF antes de salvar, evitando duplicidade de dados entre equipes.</td>
        <td>Crítica</td>
      </tr>
      <tr>
        <td><strong>RF05</strong></td>
        <td>Inteligência do Local de Votação (TRE)</td>
        <td>Seleção encadeada de Estado, Município, Zona Eleitoral, Seção e Ponto de Votação com base no banco oficial do TRE Roraima/Nacional.</td>
        <td>Alta</td>
      </tr>
      <tr>
        <td><strong>RF06</strong></td>
        <td>Sincronização Offline PWA</td>
        <td>Possibilidade de cadastrar eleitores e solicitações em áreas rurais ou sem sinal. Os dados são salvos em fila criptografada local e sincronizados ao reestabelecer conexão.</td>
        <td>Crítica</td>
      </tr>
      <tr>
        <td><strong>RF07</strong></td>
        <td>Gestão de Vouchers de Combustível</td>
        <td>Solicitação, aprovação e emissão de autorizações de abastecimento para equipes de carreata/campo com histórico e limites operacionais.</td>
        <td>Alta</td>
      </tr>
      <tr>
        <td><strong>RF08</strong></td>
        <td>Solicitação de Materiais de Campanha</td>
        <td>Pedagógico de santinhos, bandeiras, praguinhas e camisetas com fluxo de aprovação e protocolo de entrega.</td>
        <td>Média</td>
      </tr>
      <tr>
        <td><strong>RF09</strong></td>
        <td>Ouvidoria e Registro de Demandas</td>
        <td>Registro de solicitações comunitárias (infraestrutura, apoio local, reclamações) com status de acompanhamento pela coordenação central.</td>
        <td>Média</td>
      </tr>
      <tr>
        <td><strong>RF10</strong></td>
        <td>Monitor de Agendas e Missões Locais</td>
        <td>Proposição e confirmação de eventos, caminhadas e reuniões do candidato com geolocalização e confirmação de presença.</td>
        <td>Alta</td>
      </tr>
      <tr>
        <td><strong>RF11</strong></td>
        <td>Link Público de Auto-Cadastro de Eleitor</td>
        <td>URL parametrizada para que eleitores façam o próprio cadastro, alocando-os automaticamente na equipe do líder que compartilhou.</td>
        <td>Alta</td>
      </tr>
      <tr>
        <td><strong>RF12</strong></td>
        <td>Relatórios e Exportação Multi-formato</td>
        <td>Exportação instantânea de listas filtradas em Excel (.xlsx) e formulários individuais de cadastro em formato impresso/PDF.</td>
        <td>Média</td>
      </tr>
    </tbody>
  </table>

  <h2>4. REQUISITOS NÃO-FUNCIONAIS (RNF)</h2>
  <ul>
    <li><strong>RNF01 - Desempenho e Tempo de Resposta:</strong> O tempo de carregamento inicial (First Contentful Paint) é otimizado via lazy-loading de componentes, mantendo rendering &lt; 1.5s em redes 4G/3G.</li>
    <li><strong>RNF02 - Usabilidade e Acessibilidade (UI/UX):</strong> Design responsivo com regras rígidas de contraste (WCAG AA), tipografia legível em telas pequenas sob luz solar e suporte aos modos Claro e Escuro.</li>
    <li><strong>RNF03 - Segurança e Privacidade (LGPD):</strong> Os dados de eleitores são protegidos por Firestore Security Rules baseadas em token JWT de autenticação. NENHUMA chave privada é exposta no cliente.</li>
    <li><strong>RNF04 - Disponibilidade &amp; Tolerância a Falhas:</strong> Operação contínua garantida pelo armazenamento local temporário ('IndexedDB'), evitando perda de trabalho do militante em campo caso a bateria ou o sinal oscilem.</li>
    <li><strong>RNF05 - Escalabilidade:</strong> Estrutura NoSQL Firestore projetada para suportar centenas de milhares de registros por município sem degradação de performance por meio de paginação reativa.</li>
  </ul>

  <h2>5. FLUXO DE DADOS E HIERARQUIA OPERACIONAL</h2>
  <div class="highlight-box">
    <p><strong>Cadeia de Comando Tático:</strong></p>
    <p><strong>Candidato / Coordenação Geral</strong> &rarr; Define metas regionais, aprova combustível/materiais e analisa o mapa de calor de votos.</p>
    <p><strong>Coordenador Regional</strong> &rarr; Gerencia os líderes do seu município/zona, valida agendas locais e fiscaliza a expansão de base.</p>
    <p><strong>Líder de Campo (Cabo Eleitoral)</strong> &rarr; Cadastra eleitores em campo, solicita suporte logístico e engaja sua rede de vizinhos.</p>
    <p><strong>Eleitor</strong> &rarr; Confirma apoio, cadastra demandas de bairro e acompanha ações da campanha.</p>
  </div>

  <h2>6. CONCLUSÃO E HOMOLOGAÇÃO</h2>
  <p>
    A arquitetura descrita atende rigorosamente aos padrões modernos de engenharia de software para sistemas críticos de campanha política, oferecendo robustez, conformidade legal, rapidez e total autonomia operacional para equipes de rua e centro de comando.
  </p>

  <div class="footer-text">
    <p>Nexus Política 2026 &copy; Todos os direitos reservados. Documento gerado automaticamente pela plataforma.</p>
  </div>

</body>
</html>
  `.trim();

  const blob = new Blob(['\ufeff' + documentContent], {
    type: 'application/msword;charset=utf-8'
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'Arquitetura_e_Requisitos_Nexus_Politica_2026.doc';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export function downloadStrategicPlanDoc() {
  const documentContent = `
<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <meta charset="utf-8">
  <title>Plano Estratégico de Negócio - Nexus Política & Eagle Intelligence 2026</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    body {
      font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.5;
      color: #1a1a1a;
      margin: 40pt;
    }
    h1 {
      font-size: 20pt;
      color: #1e3a8a;
      border-bottom: 2px solid #1e3a8a;
      padding-bottom: 6px;
      margin-top: 24pt;
      margin-bottom: 12pt;
      text-transform: uppercase;
    }
    h2 {
      font-size: 14pt;
      color: #2563eb;
      margin-top: 18pt;
      margin-bottom: 8pt;
      border-bottom: 1px solid #e5e7eb;
      padding-bottom: 4px;
    }
    h3 {
      font-size: 12pt;
      color: #1f2937;
      margin-top: 12pt;
      margin-bottom: 6pt;
    }
    p, li {
      font-size: 11pt;
      text-align: justify;
    }
    ul, ol {
      margin-top: 4pt;
      margin-bottom: 8pt;
      padding-left: 20pt;
    }
    li {
      margin-bottom: 4pt;
    }
    .header-box {
      background-color: #f8fafc;
      border: 1px solid #cbd5e1;
      padding: 16pt;
      border-radius: 4px;
      margin-bottom: 20pt;
    }
    .header-box h1 {
      margin-top: 0;
      color: #0f172a;
      border-bottom: none;
      padding-bottom: 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10pt;
      margin-bottom: 15pt;
    }
    th, td {
      border: 1px solid #cbd5e1;
      padding: 8pt 10pt;
      text-align: left;
      font-size: 10pt;
    }
    th {
      background-color: #f1f5f9;
      color: #0f172a;
      font-weight: bold;
    }
    tr:nth-child(even) {
      background-color: #f8fafc;
    }
    .highlight-box {
      background-color: #eff6ff;
      border-left: 4px solid #2563eb;
      padding: 10pt 12pt;
      margin-top: 10pt;
      margin-bottom: 10pt;
    }
    .badge {
      background-color: #dbeafe;
      color: #1e40af;
      padding: 2px 6px;
      border-radius: 3px;
      font-weight: bold;
      font-size: 9pt;
    }
    .footer-text {
      margin-top: 30pt;
      font-size: 9pt;
      color: #64748b;
      border-top: 1px solid #e2e8f0;
      padding-top: 8pt;
      text-align: center;
    }
  </style>
</head>
<body>

  <div class="header-box">
    <h1>PLANO ESTRATÉGICO DE NEGÓCIO</h1>
    <p><strong>Plataforma:</strong> Nexus Política 2026 (Eagle Intelligence Systems)</p>
    <p><strong>Mercado Alvo:</strong> Campanhas Eleitorais Municipais, Estaduais e Federais no Brasil</p>
    <p><strong>Versão:</strong> 1.0 (Visão 2026-2028)</p>
    <p><strong>Status:</strong> Documento de Diretrizes Estratégicas e Expansão Comercial</p>
  </div>

  <h2>1. RESUMO EXECUTIVO E VISÃO DE MERCADO</h2>
  <p>
    O <strong>Nexus Política 2026</strong> é uma plataforma GovTech / PoliticalTech desenvolvida para transformar a gestão tática de campanhas eleitorais no Brasil. O sistema substitui planilhas descentralizadas e comunicação informal por uma infraestrutura centralizada de inteligência eleitoral, conectando em tempo real candidatos, coordenadores gerais, coordenadores municipais e cabos eleitorais (militância de campo).
  </p>

  <h2>2. ANÁLISE DO PROBLEMA &amp; OPORTUNIDADE</h2>
  <ul>
    <li><strong>Desorganização de Base:</strong> Perda de votos por falta de acompanhamento direto dos eleitores mapeados pelos cabos eleitorais.</li>
    <li><strong>Desperdício de Recursos:</strong> Distribuição ineficiente de materiais e vouchers de combustível sem comprovação geográfica (comprovação georreferenciada).</li>
    <li><strong>Opacidade Tática:</strong> Falta de visibilidade da liderança sobre a densidade de votos por zona eleitoral (TRE/TSE).</li>
    <li><strong>Oportunidade Comercial:</strong> Mais de 5.500 municípios e milhares de candidaturas estaduais e federais com orçamento concentrado no ciclo eleitoral de 2026.</li>
  </ul>

  <h2>3. MODELO DE NEGÓCIO E PRECIFICAÇÃO (REVENUE MODEL)</h2>
  <table>
    <thead>
      <tr>
        <th>Plano / Modalidade</th>
        <th>Público-Alvo</th>
        <th>Precificação Média</th>
        <th>Recursos Incluídos</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>Plano Municipal (Vereador/Prefeito)</strong></td>
        <td>Candidatos locais</td>
        <td>R$ 3.500,00 – R$ 8.000,00 / campanha</td>
        <td>Até 50 cabos eleitorais, mapa de calor, controle de vales e aplicativo PWA offline.</td>
      </tr>
      <tr>
        <td><strong>Plano Estadual (Dep. Estadual/Federal)</strong></td>
        <td>Candidatos de médio porte</td>
        <td>R$ 15.000,00 – R$ 35.000,00 / campanha</td>
        <td>Cabos ilimitados, geolocalização TRE, múltiplos coordenadores, relatórios em PDF/DOC.</td>
      </tr>
      <tr>
        <td><strong>Plano Majoritário (Senado/Governador)</strong></td>
        <td>Grandes coligações</td>
        <td>R$ 50.000,00 – R$ 120.000,00 / campanha</td>
        <td>Infraestrutura isolada, suporte dedicado, WhatsApp automatizado e inteligência eleitoral.</td>
      </tr>
    </tbody>
  </table>

  <h2>4. ESTRATÉGIA DE GO-TO-MARKET (GTM) E VENDAS</h2>
  <ol>
    <li><strong>Prospecção Direta (Outbound B2G/B2Pol):</strong> Abordagem direta a partidos políticos, diretórios regionais e consultores de marketing político.</li>
    <li><strong>Demonstração Interativa (Demo Mode):</strong> Uso do módulo "Demo" integrado na própria plataforma, permitindo ao coordenador experimentar os papéis antes de fechar contrato.</li>
    <li><strong>Parcerias Estratégicas:</strong> Conexão com agências de comunicação política e escritórios de advocacia eleitoral.</li>
  </ol>

  <h2>5. PLANO DE AÇÃO E MARCOS TEMPORAIS (ROADMAP 2026)</h2>
  <ul>
    <li><strong>Fase 1 (Pré-Campanha):</strong> Lançamento do MVP, captação das primeiras 20 campanhas piloto e validação das integrações de dados TRE.</li>
    <li><strong>Fase 2 (Convenções Partidárias):</strong> Expansão comercial e onboarding massivo das equipes de campo (cabos eleitorais).</li>
    <li><strong>Fase 3 (Período Eleitoral Ativo):</strong> Monitoramento intensivo de disparos, gestão de vales de combustível e relatórios diários de meta.</li>
    <li><strong>Fase 4 (Pós-Eleição):</strong> Transição dos dados para mandato legislativo/pós-venda para retenção do candidato.</li>
  </ul>

  <h2>6. ANÁLISE DE RISCOS E CONFORMIDADE (LGPD ELEITORAL)</h2>
  <div class="highlight-box">
    <p><strong>Garantias de Privacidade e Segurança:</strong></p>
    <p>• Dados armazenados com encriptação e isolamento por candidato.</p>
    <p>• Consentimento explícito dos eleitores cadastrados conforme a LGPD.</p>
    <p>• Logs auditáveis de ações de cabos eleitorais e coordenadores.</p>
  </div>

  <div class="footer-text">
    <p>Nexus Política 2026 &copy; Eagle Intelligence Systems. Documento Confidencial de Planejamento Estratégico.</p>
  </div>

</body>
</html>
  `.trim();

  const blob = new Blob(['\ufeff' + documentContent], {
    type: 'application/msword;charset=utf-8'
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'Plano_Estrategico_Nexus_Politica_2026.doc';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

