// RNP Project Staffing Reference (Core Team and Professional Profiles) — English and Brazilian Portuguese editions
//   Source of truth: Chinese baseline v2 of 2026-09-15 (build_rnp_team_proposal_docx.js): HQ R&D headcounts of 15 Sep 20:52,
//   English as the only working language. Same structure, same numbers (core 35 / extended 131 / total 166, 8 groups, 26 roles); professional translation.
//   node build_rnp_team_proposal_i18n.js en   ->  RNP-Project-Staffing-Reference-Core-Team-and-Professional-Profiles-EN-20260915.docx
//   node build_rnp_team_proposal_i18n.js pt   ->  Referencia-de-Composicao-da-Equipe-da-RNP-Equipe-Nucleo-e-Perfis-Profissionais-PT-BR-20260915.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
        ShadingType, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, TabStopType } = require("docx");

const LANG = (process.argv[2] || "en").toLowerCase();
if (!["en", "pt"].includes(LANG)) { console.error("usage: node build_rnp_team_proposal_i18n.js en|pt"); process.exit(1); }
const L = s => (typeof s === "string" ? s : s[LANG]);

const F = "Calibri";
const RED = "C00000", NAVY = "1F3A5F", BLUE = "0057A8", GRAY = "6B6B6B", BLACK = "1D1D1A", LIGHT = "F3F6FA", LINE = "C9D3E0";
const PAGE_W = 9638;
const D = [];

// =====================================================================================
// UI strings
// =====================================================================================
const UI = {
  title:    { en: "RNP Project Staffing Reference", pt: "Referência de Composição da Equipe da RNP para o Projeto" },
  subtitle: { en: "Core Team and Professional Profiles", pt: "Equipe Núcleo e Perfis Profissionais" },
  meta:     { en: "Submitted by iFLYTEK for RNP's reference  ·  15 September 2026", pt: "Apresentado pela iFLYTEK para referência da RNP  ·  15 de setembro de 2026" },
  intro1:   { en: "Based on the scope of work of this project, the practical needs of RNP's side in receiving the technology transfer and implementing the project jointly, and iFLYTEK's experience from similar projects, iFLYTEK presents this staffing reference for RNP's side: the roles required, the competency requirements for each role, the suggested headcount, and how each role works with iFLYTEK.",
              pt: "Com base no escopo de trabalho deste projeto, nas necessidades práticas do lado da RNP para receber a transferência de tecnologia e implementar o projeto em conjunto, e na experiência da iFLYTEK em projetos semelhantes, a iFLYTEK apresenta esta referência de composição da equipe da RNP: os cargos necessários, os requisitos de competência de cada cargo, o quantitativo sugerido e a forma de colaboração de cada cargo com a iFLYTEK." },
  intro2:   { en: "Headcounts are estimated from the current division of the scope of work and the workload of each phase, and are offered as a reference for RNP's team planning. iFLYTEK will be glad to review the plan jointly once RNP has defined its team, and to refine the knowledge transfer arrangements accordingly.",
              pt: "Os quantitativos foram estimados a partir da divisão atual do escopo de trabalho e da carga de trabalho de cada fase, e são oferecidos como referência para o planejamento da equipe da RNP. A iFLYTEK terá satisfação em revisar o plano em conjunto assim que a RNP definir sua equipe, e em detalhar os arranjos de transferência de conhecimento em conformidade." },
  s1:       { en: "1.  Suggested team structure", pt: "1.  Estrutura sugerida da equipe" },
  s1p0:     { en: "It is suggested that RNP's team be organised in two tiers, totalling {T} people in {G} groups and {R} roles.",
              pt: "Sugere-se que a equipe da RNP seja organizada em duas camadas, totalizando {T} pessoas em {G} grupos e {R} cargos." },
  s1core:   { en: "**Core team (Equipe núcleo): {C} people.** A long-term team present throughout the project, the main audience of the knowledge transfer, and the team that continues to run the platform, models and application after the project ends. It comprises three kinds of roles: those accountable for project deliverables and for signing acceptance; those that will run the platform, models and application independently after the project; and the technical backbone that receives the knowledge transfer and passes it on to others.",
              pt: "**Equipe núcleo: {C} pessoas.** Equipe de longo prazo, presente ao longo de todo o projeto, principal público da transferência de conhecimento e responsável por continuar operando a plataforma, os modelos e a aplicação após o término do projeto. Compreende três tipos de cargos: os responsáveis pelos resultados do projeto e pela assinatura dos aceites; os que operarão de forma independente a plataforma, os modelos e a aplicação após o projeto; e o núcleo técnico que recebe a transferência de conhecimento e a repassa aos demais." },
  s1ext:    { en: "**Extended team (Equipe complementar): {E} people.** Carries out phase-based, workload-intensive tasks such as corpus production, domain and language expertise, and application localisation development and operations, joining as the corresponding work phases begin. It is suggested that each type of extended-team role be paired with a core-team role as its point of contact, to ensure continuity of work and retention of results. Extended-team positions may be staffed through research projects, fixed-term contracts and the participation of universities and partner institutions: workload-intensive positions such as corpus annotation are well suited to university students and temporary staff, while language and domain experts are well suited to experts from universities and specialist institutions.",
              pt: "**Equipe complementar: {E} pessoas.** Executa tarefas por fase e de grande volume de trabalho, como produção de corpus, apoio especializado de domínio e de língua, e desenvolvimento e operação da localização da aplicação, ingressando conforme as respectivas fases de trabalho. Sugere-se que cada tipo de cargo da equipe complementar tenha um cargo da equipe núcleo como ponto de contato, para garantir a continuidade do trabalho e a consolidação dos resultados. As posições da equipe complementar podem ser preenchidas por meio de projetos de pesquisa, contratos por prazo determinado e da participação de universidades e instituições parceiras: posições de grande volume de trabalho, como a anotação de corpus, são adequadas a estudantes universitários e pessoal temporário, enquanto os especialistas em língua e de domínio são adequados a especialistas de universidades e instituições especializadas." },
  s1last:   { en: "Headcounts are suggested allocations estimated from the work content; one person may cover more than one role, and the numbers may be adjusted to the actual workload of each phase.",
              pt: "Os quantitativos são alocações sugeridas, estimadas a partir do conteúdo do trabalho; uma mesma pessoa pode cobrir mais de um cargo, e os números podem ser ajustados à carga de trabalho real de cada fase." },
  s2:       { en: "2.  Overview of roles and headcount", pt: "2.  Visão geral dos cargos e quantitativos" },
  hdr:      { en: ["No.", "Group", "Role", "Core", "Ext.", "Summary"], pt: ["Nº", "Grupo", "Cargo", "Núcleo", "Compl.", "Resumo"] },
  core:     { en: "Core", pt: "Núcleo" }, ext: { en: "Ext.", pt: "Compl." },
  total:    { en: "Total", pt: "Total" },
  totalLine:{ en: "Core team {C} + extended team {E} = {T} people", pt: "Equipe núcleo {C} + equipe complementar {E} = {T} pessoas" },
  s3:       { en: "3.  Professional profiles", pt: "3.  Perfis profissionais" },
  s3a:      { en: "Each profile has two parts:", pt: "Cada perfil tem duas partes:" },
  s3b:      { en: "**Requirements** describe the kind of person the role needs;", pt: "**Requisitos** descrevem o tipo de profissional de que o cargo precisa;" },
  s3c:      { en: "**Collaboration** describes which iFLYTEK team the role works with in the project and how. Numbering follows the overview table; the suggested core and extended headcounts appear at the right of each heading.",
              pt: "**Colaboração** descreve com qual equipe da iFLYTEK o cargo trabalha no projeto e de que forma. A numeração segue a tabela de visão geral; os quantitativos sugeridos de núcleo e complementar aparecem à direita de cada título." },
  req:      { en: "Requirements", pt: "Requisitos" }, col: { en: "Collaboration", pt: "Colaboração" },
  footer:   { en: "iFLYTEK  ·  RNP Project Staffing Reference  ·  September 2026  ·  Page ", pt: "iFLYTEK  ·  Referência de Composição da Equipe da RNP  ·  Setembro de 2026  ·  Página " },
  out:      { en: "RNP-Project-Staffing-Reference-Core-Team-and-Professional-Profiles-EN-20260915.docx", pt: "Referencia-de-Composicao-da-Equipe-da-RNP-Equipe-Nucleo-e-Perfis-Profissionais-PT-BR-20260915.docx" },
};

// =====================================================================================
// Content: 8 groups, 26 roles (numbers identical to the Chinese baseline)
// =====================================================================================
const GROUPS = [
  { name: { en: "Project and Product", pt: "Projeto e Produto" },
    intro: { en: "Responsible for organising the project on RNP's side, including schedule, resources and acceptance, and for defining the National AI Application as a product. Work starts in project phase 1.",
             pt: "Responsável pela organização do projeto do lado da RNP, incluindo cronograma, recursos e aceite, e pela definição da Aplicação Nacional de IA como produto. O trabalho começa na fase 1 do projeto." },
    roles: [
      { name: { en: "Project Manager", pt: "Gerente de Projeto" }, en: "Project Manager", c: 1, e: 1,
        one: { en: "Coordinates RNP-side schedule, resources and acceptance; main counterpart for iFLYTEK and Huawei", pt: "Coordena cronograma, recursos e aceite do lado da RNP; interlocutor principal com a iFLYTEK e a Huawei" },
        note: { en: "It is suggested that one of the two also coordinates knowledge transfer and training.", pt: "Sugere-se que um dos dois também coordene a transferência de conhecimento e o treinamento." },
        req: { en: ["Delivery management experience on large IT or infrastructure projects; able to coordinate RNP, iFLYTEK, Huawei and the ministry", "Familiar with the contract and acceptance procedures of government projects", "Able to work in English"],
               pt: ["Experiência em gestão de entrega de grandes projetos de TI ou infraestrutura; capaz de coordenar RNP, iFLYTEK, Huawei e o ministério", "Familiaridade com os procedimentos contratuais e de aceite de projetos governamentais", "Capacidade de trabalhar em inglês"] },
        col: { en: ["Counterpart of iFLYTEK's project team: jointly maintains the master schedule and milestones; takes part in regular meetings and phase reviews", "Organises RNP-side acceptance reviews; acceptance documents are signed by authorised personnel"],
               pt: ["Contraparte da equipe de projeto da iFLYTEK: mantém em conjunto o cronograma geral e os marcos; participa das reuniões periódicas e das revisões de fase", "Organiza as revisões de aceite do lado da RNP; os documentos de aceite são assinados por pessoal autorizado"] } },
      { name: { en: "Product Manager", pt: "Gerente de Produto" }, en: "Product Manager", c: 1, e: 1,
        one: { en: "Defines the functions, user scope and content rules of the National AI Application", pt: "Define as funcionalidades, o escopo de usuários e as regras de conteúdo da Aplicação Nacional de IA" },
        req: { en: ["Experience with internet products for the public or for government services", "Able to turn Brazilian local content and use cases into product requirements", "Understands the capabilities and limits of LLM-based applications"],
               pt: ["Experiência com produtos digitais voltados ao público ou a serviços governamentais", "Capaz de transformar conteúdos e casos de uso locais brasileiros em requisitos de produto", "Compreende as capacidades e os limites de aplicações baseadas em grandes modelos de linguagem"] },
        col: { en: ["Counterpart of iFLYTEK's product team: confirms user scope, content rules and configuration requirements within the agreed functional scope; new features and roadmap proposals are evaluated and confirmed separately", "Takes over application configuration and releases after shadowing"],
               pt: ["Contraparte da equipe de produto da iFLYTEK: confirma escopo de usuários, regras de conteúdo e requisitos de configuração dentro do escopo funcional acordado; novas funcionalidades e propostas de roadmap são avaliadas e confirmadas separadamente", "Assume a configuração e a publicação da aplicação após o período de acompanhamento"] } },
      { name: { en: "PMO", pt: "PMO" }, en: "PMO", c: 0, e: 2,
        one: { en: "Maintains the project plan, action items, deliverable lists and documentation; tracks issues and the follow-up of decisions", pt: "Mantém o plano do projeto, os itens de ação, as listas de entregáveis e a documentação; acompanha problemas e a execução das decisões" },
        req: { en: ["Experience in project coordination and schedule tracking", "Familiar with action-item and issue logs and document management", "Strong organisational and execution skills"],
               pt: ["Experiência em coordenação de projetos e acompanhamento de cronograma", "Familiaridade com registros de itens de ação e de problemas e com gestão documental", "Forte capacidade de organização e execução"] },
        col: { en: ["Counterpart of iFLYTEK's project management staff; maintains meeting minutes, action items and deliverable lists", "Tracks open items on both sides; coordinates reviews, acceptance and document archiving"],
               pt: ["Contraparte da equipe de gestão de projeto da iFLYTEK; mantém atas de reunião, itens de ação e listas de entregáveis", "Acompanha as pendências de ambas as partes; coordena revisões, aceites e arquivamento de documentos"] } },
      { name: { en: "Infrastructure Specialist", pt: "Especialista em Infraestrutura" }, en: "Infrastructure Specialist", c: 2, e: 0,
        one: { en: "Coordinates the resources, network, accounts and environment conditions the platform needs; organises troubleshooting of base-environment issues", pt: "Coordena recursos, rede, contas e condições de ambiente necessários à operação da plataforma; organiza a investigação de problemas do ambiente básico" },
        req: { en: ["Solid grounding in compute, storage and networking; familiar with Linux, containers and AI platform deployment environments", "Able to identify resource, permission, network and software-dependency issues", "Able to coordinate infrastructure and application teams in troubleshooting"],
               pt: ["Base sólida em computação, armazenamento e redes; familiaridade com Linux, contêineres e ambientes de implantação de plataformas de IA", "Capaz de identificar problemas de recursos, permissões, rede e dependências de software", "Capaz de coordenar as equipes de infraestrutura e de aplicação na resolução de falhas"] },
        col: { en: ["Counterpart of iFLYTEK's deployment team; confirms resource, network, account and environment requirements", "Coordinates the responsible parties to resolve base-environment issues; supports deployment checks and joint diagnosis"],
               pt: ["Contraparte da equipe de implantação da iFLYTEK; confirma os requisitos de recursos, rede, contas e ambiente", "Coordena as partes responsáveis para resolver problemas do ambiente básico; apoia verificações de implantação e diagnóstico conjunto"] } },
    ] },
  { name: { en: "Corpus and Data Resources", pt: "Corpus e Recursos de Dados" },
    intro: { en: "Organises the acquisition, licensing, production and quality control of the Portuguese corpus; the main workforce of project phase 1.",
             pt: "Organiza a aquisição, o licenciamento, a produção e o controle de qualidade do corpus em português; principal força de trabalho da fase 1 do projeto." },
    roles: [
      { name: { en: "Data Resource Engineer", pt: "Engenheiro de Recursos de Dados" }, en: "Data Resource Engineer", c: 2, e: 2,
        one: { en: "Identifies data sources, secures licences and organises corpus acquisition", pt: "Identifica fontes de dados, assegura as licenças e organiza a aquisição do corpus" },
        req: { en: ["Familiar with Brazilian public and institutional data resources: government, universities, media and research databases", "Understands data licensing; able to review source licence terms", "Experience in data collection and curation; native Portuguese speaker; able to work in English"],
               pt: ["Familiaridade com recursos de dados públicos e institucionais brasileiros: governo, universidades, mídia e bases de pesquisa", "Conhece licenciamento de dados; capaz de analisar os termos de licença das fontes", "Experiência em coleta e organização de dados; falante nativo de português; capacidade de trabalhar em inglês"] },
        col: { en: ["Counterpart of iFLYTEK's data team: confirms sources and quotas against the corpus specification; collects and registers licensing documents and submits them to RNP's legal team to confirm the licence conditions", "Through foundation courses and shadowing, masters data sources, licensing and the platform processing workflow"],
               pt: ["Contraparte da equipe de dados da iFLYTEK: confirma fontes e cotas conforme a especificação do corpus; reúne e registra os documentos de autorização e os submete ao jurídico da RNP para confirmação das condições de licença", "Por meio de cursos básicos e acompanhamento, domina as fontes de dados, o licenciamento e o fluxo de processamento na plataforma"] } },
      { name: { en: "Data Quality Engineer", pt: "Engenheiro de Qualidade de Dados" }, en: "Data Quality Engineer", c: 2, e: 2,
        one: { en: "Applies the corpus counting, sampling, quality-threshold and re-check rules", pt: "Aplica as regras de contagem, amostragem, limiares de qualidade e reverificação do corpus" },
        req: { en: ["Experience in data quality management and sample inspection", "Understands data formats, annotation and quality specifications", "Able to produce quality statistics, defect analysis and reports", "Able to work in English"],
               pt: ["Experiência em gestão da qualidade de dados e inspeção por amostragem", "Compreende formatos de dados, anotação e especificações de qualidade", "Capaz de produzir estatísticas de qualidade, análise de defeitos e relatórios", "Capacidade de trabalhar em inglês"] },
        col: { en: ["Counterpart of iFLYTEK's data team; confirms the sampling, quality-judgement and re-check rules", "Performs RNP-side quality checks and re-checks after corrections; supports iFLYTEK's agreed sample inspections and result analysis"],
               pt: ["Contraparte da equipe de dados da iFLYTEK; confirma as regras de amostragem, julgamento de qualidade e reverificação", "Executa as verificações de qualidade do lado da RNP e a reverificação após correções; apoia as inspeções por amostragem acordadas com a iFLYTEK e a análise dos resultados"] } },
      { name: { en: "Data Engineer", pt: "Engenheiro de Dados" }, en: "Data Engineer", c: 0, e: 4,
        one: { en: "Performs batch data ingestion, cleaning, transformation and de-duplication; maintains processing pipelines and data versions", pt: "Executa a ingestão em lote, limpeza, transformação e deduplicação dos dados; mantém os fluxos de processamento e as versões dos dados" },
        req: { en: ["Proficient in Python, SQL and data processing tools", "Experience with batch cleaning, transformation, de-duplication and pipelines", "Familiar with data versioning, metadata and source traceability"],
               pt: ["Proficiência em Python, SQL e ferramentas de processamento de dados", "Experiência em limpeza, transformação e deduplicação em lote e em pipelines", "Familiaridade com versionamento de dados, metadados e rastreabilidade de origem"] },
        col: { en: ["Counterpart of iFLYTEK's data and platform teams; carries out the assigned data processing tasks and reports anomalies", "Receives platform operation guidance; maintains the processing pipelines, configurations and version records for which RNP is responsible"],
               pt: ["Contraparte das equipes de dados e de plataforma da iFLYTEK; executa as tarefas de processamento de dados atribuídas e reporta anomalias", "Recebe orientação de operação da plataforma; mantém os fluxos de processamento, as configurações e os registros de versão sob responsabilidade da RNP"] } },
      { name: { en: "Annotator", pt: "Anotador" }, en: "Annotator", c: 0, e: 100,
        one: { en: "Produces text annotation, safety samples and speech transcription to the agreed specifications", pt: "Produz anotação de texto, amostras de segurança e transcrição de fala conforme as especificações acordadas" },
        req: { en: ["Native Portuguese speaker with sound written language; careful and consistent; able to work in English", "After training, able to work to the annotation specification; transcription positions require good listening and dictation skills", "May consist mainly of university students or temporary staff; keeping a stable core group is recommended"],
               pt: ["Falante nativo de português, com boa expressão escrita; cuidadoso e consistente; capacidade de trabalhar em inglês", "Após treinamento, capaz de trabalhar conforme a especificação de anotação; as posições de transcrição exigem boa capacidade de escuta e ditado", "Pode ser composto principalmente por estudantes universitários ou pessoal temporário; recomenda-se manter um grupo fixo de referência"] },
        col: { en: ["Works with the annotation specification, samples and tools confirmed by both parties; supports iFLYTEK's quality re-checks within the agreed batches, sampling scope and re-check rounds"],
               pt: ["Trabalha com a especificação de anotação, as amostras e as ferramentas confirmadas por ambas as partes; apoia as reverificações de qualidade da iFLYTEK dentro dos lotes, do escopo de amostragem e das rodadas de reverificação acordados"] } },
    ] },
  { name: { en: "Portuguese Language Experts", pt: "Especialistas em Língua Portuguesa" },
    intro: { en: "Defines the quality, annotation and speech standards for Brazilian Portuguese. One core member is suggested to lead; the other experts may come from universities and research institutions.",
             pt: "Define os padrões de qualidade, anotação e fala para o português brasileiro. Sugere-se que um membro da equipe núcleo lidere; os demais especialistas podem vir de universidades e instituições de pesquisa." },
    roles: [
      { name: { en: "Portuguese Language Expert", pt: "Especialista em Língua Portuguesa" }, en: "Portuguese Language Expert", c: 1, e: 3,
        one: { en: "Sets the Portuguese quality standards, annotation rules, safety labels and speech standards", pt: "Define os padrões de qualidade do português, as regras de anotação, os rótulos de segurança e os padrões de fala" },
        req: { en: ["Background in linguistics, computational linguistics or phonetics", "Familiar with regional variation and contemporary usage of Brazilian Portuguese", "Experience in corpus or annotation projects", "Able to work in English"],
               pt: ["Formação em linguística, linguística computacional ou fonética", "Familiaridade com a variação regional e o uso contemporâneo do português brasileiro", "Experiência em projetos de corpus ou de anotação", "Capacidade de trabalhar em inglês"] },
        col: { en: ["Counterpart of iFLYTEK's language and speech teams: jointly walks through the annotation standards and safety labels and confirms the rules", "Assesses the quality of the model's Portuguese output"],
               pt: ["Contraparte das equipes de linguagem e de fala da iFLYTEK: revisa em conjunto os padrões de anotação e os rótulos de segurança e confirma as regras", "Avalia a qualidade das saídas do modelo em português"] } },
    ] },
  { name: { en: "Compliance and Review", pt: "Conformidade e Revisão" },
    intro: { en: "Responsible for content review and for legal compliance relating to data, models and the application; takes over the Content Safety Review Platform after acceptance. Operation of security equipment and technical alerts are the responsibility of the security operations role.",
             pt: "Responsável pela revisão de conteúdo e pela conformidade legal relativa a dados, modelos e aplicação; assume a Plataforma de Revisão de Segurança de Conteúdo após o aceite. A operação dos equipamentos de segurança e os alertas técnicos são de responsabilidade do cargo de operações de segurança." },
    roles: [
      { name: { en: "Content Review Specialist", pt: "Especialista em Revisão de Conteúdo" }, en: "Content Review Specialist", c: 1, e: 2,
        one: { en: "Maintains risk labels, lexicons and review policies; performs manual re-checks and false-positive analysis", pt: "Mantém rótulos de risco, léxicos e políticas de revisão; realiza reverificações manuais e análise de falsos positivos" },
        req: { en: ["Experience in content safety or information review", "Understands the boundaries of sensitive content in the Brazilian legal and social context", "Able to define and continuously maintain review rules"],
               pt: ["Experiência em segurança de conteúdo ou revisão de informações", "Compreende os limites de conteúdo sensível no contexto legal e social brasileiro", "Capaz de definir e manter continuamente as regras de revisão"] },
        col: { en: ["Counterpart of iFLYTEK's safety team: jointly configures policies, lexicons and samples on the Content Safety Review Platform", "Operates the platform independently after acceptance"],
               pt: ["Contraparte da equipe de segurança da iFLYTEK: configura em conjunto políticas, léxicos e amostras na Plataforma de Revisão de Segurança de Conteúdo", "Opera a plataforma de forma independente após o aceite"] } },
      { name: { en: "Legal Counsel", pt: "Assessor Jurídico" }, en: "Legal Counsel", c: 2, e: 0,
        one: { en: "Secures data licensing and compliance; reviews the legal risks of the corpus, models and application", pt: "Assegura o licenciamento e a conformidade dos dados; analisa os riscos jurídicos do corpus, dos modelos e da aplicação" },
        req: { en: ["Legal background in data protection (LGPD), intellectual property and licensing", "Able to review data source licences and open-source licence terms", "Experience in public procurement and contracts preferred"],
               pt: ["Formação jurídica em proteção de dados (LGPD), propriedade intelectual e licenciamento", "Capaz de analisar licenças de fontes de dados e termos de licenças de código aberto", "Experiência em contratações públicas e contratos é um diferencial"] },
        col: { en: ["Counterpart of iFLYTEK's data and commercial teams: confirms each data source licence; reviews the licence terms for model weights, the platform and the application"],
               pt: ["Contraparte das equipes de dados e comercial da iFLYTEK: confirma cada licença de fonte de dados; analisa os termos de licença dos pesos dos modelos, da plataforma e da aplicação"] } },
      { name: { en: "Review Lead", pt: "Líder de Revisão" }, en: "Review Lead", c: 1, e: 0,
        one: { en: "Organises confirmation of review rules, quality re-checks and escalation of disputes; coordinates the people involved to resolve cases", pt: "Organiza a confirmação das regras de revisão, as reverificações de qualidade e o escalonamento de controvérsias; coordena os envolvidos na resolução dos casos" },
        req: { en: ["Experience managing compliance or review teams", "Able to define processes, record-keeping and escalation mechanisms", "Able to work with legal, business and technical teams"],
               pt: ["Experiência na gestão de equipes de conformidade ou revisão", "Capaz de definir processos, registros e mecanismos de escalonamento", "Capaz de trabalhar com as equipes jurídica, de negócio e técnica"] },
        col: { en: ["Agrees the review escalation and approval rules with iFLYTEK; RNP-side owner for acceptance and handover of the review platform"],
               pt: ["Acorda com a iFLYTEK as regras de escalonamento e aprovação da revisão; responsável do lado da RNP pelo aceite e pela transferência da plataforma de revisão"] } },
    ] },
  { name: { en: "Business and Domain Experts", pt: "Especialistas de Negócio e de Domínio" },
    intro: { en: "Defines the business scenarios for the three domain models (DLMs) and assesses domain results. Work starts in project phase 2.",
             pt: "Define os cenários de negócio dos três modelos de domínio (DLM) e avalia os resultados de domínio. O trabalho começa na fase 2 do projeto." },
    roles: [
      { name: { en: "Business Requirements Lead", pt: "Líder de Requisitos de Negócio" }, en: "Business Requirements Lead", c: 3, e: 0,
        one: { en: "Defines the business scenario, inputs and acceptance for each domain model service", pt: "Define o cenário de negócio, as entradas e o aceite de cada serviço de modelo de domínio" },
        note: { en: "It is suggested that each of the three leads one domain, as RNP-side owner of that domain model.", pt: "Sugere-se que cada um dos três lidere um domínio, como responsável do lado da RNP por aquele modelo de domínio." },
        req: { en: ["Familiar with the business processes of the domain", "Able to turn business problems into scenarios a model can deliver", "Able to engage the domain institutions"],
               pt: ["Familiaridade com os processos de negócio do domínio", "Capaz de traduzir problemas de negócio em cenários que um modelo possa entregar", "Capaz de mobilizar as instituições do domínio"] },
        col: { en: ["Counterpart of iFLYTEK's domain model advisors: for each domain model, jointly fixes one scenario, the input data and the evaluation method, keeping RNP's business validation of the model separate from the acceptance of iFLYTEK's training advisory service"],
               pt: ["Contraparte dos consultores de modelos de domínio da iFLYTEK: para cada modelo de domínio, fixa em conjunto um cenário, os dados de entrada e o método de avaliação, distinguindo a validação de negócio do modelo pela RNP do aceite do serviço de assessoria de treinamento da iFLYTEK"] } },
      { name: { en: "Domain Requirements Specialist", pt: "Especialista em Requisitos de Domínio" }, en: "Domain Requirements Specialist", c: 0, e: 2,
        one: { en: "Specifies domain data and expected results; sets content boundaries", pt: "Especifica os dados de domínio e os resultados esperados; define os limites de conteúdo" },
        req: { en: ["Background in the domain's data and knowledge", "Able to define evaluation sets and expected outputs and to read evaluation reports"],
               pt: ["Formação nos dados e no conhecimento do domínio", "Capaz de definir conjuntos de avaliação e saídas esperadas e de interpretar relatórios de avaliação"] },
        col: { en: ["Counterpart of iFLYTEK's domain model advisors: jointly defines evaluation methods and content boundaries; takes part in evaluation exercises"],
               pt: ["Contraparte dos consultores de modelos de domínio da iFLYTEK: define em conjunto os métodos de avaliação e os limites de conteúdo; participa dos exercícios de avaliação"] } },
      { name: { en: "Domain Expert", pt: "Especialista de Domínio" }, en: "Domain Expert", c: 0, e: 2,
        one: { en: "Assesses model outputs within the domain", pt: "Avalia as saídas do modelo dentro do domínio" },
        req: { en: ["Senior professional in the domain; may take part on a part-time basis from universities or specialist institutions", "Able to judge the correctness and applicability of model outputs"],
               pt: ["Profissional sênior do domínio; pode participar em tempo parcial, vindo de universidades ou instituições especializadas", "Capaz de julgar a correção e a aplicabilidade das saídas do modelo"] },
        col: { en: ["Takes part in the domain evaluations organised by RNP, gives expert assessments and supports iFLYTEK in analysing the evaluation results"],
               pt: ["Participa das avaliações de domínio organizadas pela RNP, emite pareceres especializados e apoia a iFLYTEK na análise dos resultados"] } },
    ] },
  { name: { en: "Model and Test Engineering", pt: "Engenharia de Modelos e Testes" },
    intro: { en: "Takes part in the training and evaluation of the national model in the Model Training Center through technical exchange and agreed hands-on work, and carries out the training and evaluation of the domain models for which RNP is responsible; the contractual division of responsibility for national model training is unchanged. Work starts in project phase 2; this group is the main trainee group for the knowledge transfer.",
             pt: "Participa do treinamento e da avaliação do modelo nacional no Centro de Treinamento de Modelos, por meio de intercâmbio técnico e prática acordada, e executa o treinamento e a avaliação dos modelos de domínio sob responsabilidade da RNP; a divisão contratual de responsabilidades pelo treinamento do modelo nacional permanece inalterada. O trabalho começa na fase 2 do projeto; este grupo é o principal público da transferência de conhecimento." },
    roles: [
      { name: { en: "NLP/DLM Algorithm Engineer", pt: "Engenheiro de Algoritmos de PLN/DLM" }, en: "NLP/DLM Algorithm Engineer", c: 3, e: 0,
        one: { en: "Carries out the domain model training and experiments for which RNP is responsible; takes part in technical exchange and performance analysis on the national model", pt: "Executa o treinamento e os experimentos dos modelos de domínio sob responsabilidade da RNP; participa do intercâmbio técnico e da análise de desempenho do modelo nacional" },
        note: { en: "It is suggested that one or two of them take part in the continual pre-training phase carried out in China.", pt: "Sugere-se que um ou dois deles participem da fase de pré-treinamento contínuo realizada na China." },
        req: { en: ["Programming background in Python, C++ or Java, Python preferred", "Background in at least one of machine learning, natural language processing or large language models", "Understands large-model technologies such as attention, GPT and Transformer architectures", "Proficient with deep learning and LLM frameworks such as Transformers, PyTorch and Megatron", "Hands-on experience in LLM pre-training or post-training preferred"],
               pt: ["Base de programação em Python, C++ ou Java, preferencialmente Python", "Formação em pelo menos uma das áreas: aprendizado de máquina, processamento de linguagem natural ou grandes modelos de linguagem", "Compreende tecnologias de grandes modelos, como attention, GPT e arquiteturas Transformer", "Proficiência em frameworks de aprendizado profundo e de LLM, como Transformers, PyTorch e Megatron", "Experiência prática em pré-treinamento ou pós-treinamento de LLMs é um diferencial"] },
        col: { en: ["Counterpart of iFLYTEK's model team: takes part in national model training and evaluation through shadowing, and submits training jobs independently in guided hands-on sessions", "Domain models are trained by RNP with iFLYTEK's advisory support"],
               pt: ["Contraparte da equipe de modelos da iFLYTEK: participa do treinamento e da avaliação do modelo nacional em regime de acompanhamento e submete jobs de treinamento de forma autônoma nas sessões de prática assistida", "Os modelos de domínio são treinados pela RNP com apoio consultivo da iFLYTEK"] } },
      { name: { en: "ASR/TTS Algorithm Engineer", pt: "Engenheiro de Algoritmos de ASR/TTS" }, en: "ASR/TTS Algorithm Engineer", c: 1, e: 1,
        one: { en: "Checks speech data, tests models and localises issues; supports validation of the speech capabilities", pt: "Verifica dados de fala, testa modelos e localiza problemas; apoia a validação das capacidades de fala" },
        req: { en: ["Programming background in Python or C++", "Audio processing background; familiar with common audio formats, sampling rates, noise reduction and segmentation", "Familiar with deep learning frameworks such as PyTorch; understands ASR or TTS principles and evaluation methods, including common metrics such as WER/CER, MOS and RTF", "Familiarity with speech tools and models such as Whisper, ESPnet, FunASR, VITS, FireRedTTS and IndexTTS preferred", "Hands-on experience in ASR/TTS training, tuning or deployment preferred"],
               pt: ["Base de programação em Python ou C++", "Formação em processamento de áudio; familiaridade com formatos de áudio comuns, taxas de amostragem, redução de ruído e segmentação", "Familiaridade com frameworks de aprendizado profundo como PyTorch; compreende os princípios e métodos de avaliação de ASR ou TTS, incluindo métricas comuns como WER/CER, MOS e RTF", "Familiaridade com ferramentas e modelos de fala como Whisper, ESPnet, FunASR, VITS, FireRedTTS e IndexTTS é um diferencial", "Experiência prática em treinamento, ajuste ou implantação de ASR/TTS é um diferencial"] },
        col: { en: ["Counterpart of iFLYTEK's speech model team; supports data checks, service calls and performance validation", "Provides error samples and reproduction details; tracks the handling of speech issues and regression results"],
               pt: ["Contraparte da equipe de modelos de fala da iFLYTEK; apoia verificações de dados, chamadas de serviço e validação de desempenho", "Fornece amostras de erro e informações de reprodução; acompanha o tratamento dos problemas de fala e os resultados de regressão"] } },
      { name: { en: "Training and Inference Engineer", pt: "Engenheiro de Treinamento e Inferência" }, en: "Training and Inference Engineer", c: 1, e: 1,
        one: { en: "Runs training jobs and inference validation from the engineering side; handles resource, recovery and runtime issues", pt: "Executa, do ponto de vista de engenharia, os jobs de treinamento e a validação de inferência; trata problemas de recursos, recuperação e execução" },
        req: { en: ["Programming background in Python, Shell or C++", "Familiar with Linux, containers (Docker/Kubernetes) and GPU cluster environments", "Familiarity with frameworks such as PyTorch, Transformers, Megatron, DeepSpeed, vLLM, SGLang, TensorRT-LLM and Triton preferred", "Familiar with distributed training, resource scheduling, checkpoint recovery, job resubmission, monitoring and alerting, and log analysis", "Familiarity with training and inference acceleration techniques preferred, including but not limited to DP/EP/CP/PP parallelism, prefill-decode disaggregation, PagedAttention, KV cache management, speculative decoding and quantization", "Experience in large-scale distributed training, inference deployment, cluster operations or AI platforms preferred"],
               pt: ["Base de programação em Python, Shell ou C++", "Familiaridade com Linux, contêineres (Docker/Kubernetes) e ambientes de cluster de GPU", "Familiaridade com frameworks como PyTorch, Transformers, Megatron, DeepSpeed, vLLM, SGLang, TensorRT-LLM e Triton é um diferencial", "Familiaridade com treinamento distribuído, escalonamento de recursos, recuperação de checkpoints, ressubmissão de jobs, monitoramento e alertas e análise de logs", "Familiaridade com técnicas de aceleração de treinamento e inferência é um diferencial, incluindo, entre outras, paralelismo DP/EP/CP/PP, desagregação prefill-decode, PagedAttention, gestão de KV cache, speculative decoding e quantização", "Experiência em treinamento distribuído em larga escala, implantação de inferência, operação de clusters ou plataformas de IA é um diferencial"] },
        col: { en: ["Counterpart of iFLYTEK's model and platform teams; takes part in job configuration, resource checks and training and inference validation", "Performs the engineering operations for which RNP is responsible; provides logs and environment details and supports the diagnosis of complex issues"],
               pt: ["Contraparte das equipes de modelos e de plataforma da iFLYTEK; participa da configuração de jobs, das verificações de recursos e da validação de treinamento e inferência", "Executa as operações de engenharia sob responsabilidade da RNP; fornece logs e informações de ambiente e apoia o diagnóstico de problemas complexos"] } },
      { name: { en: "Model Evaluation Engineer", pt: "Engenheiro de Avaliação de Modelos" }, en: "Model Evaluation Engineer", c: 1, e: 1,
        one: { en: "Manages independent evaluation sets; runs performance, safety and regression tests and produces verifiable results", pt: "Gerencia conjuntos de avaliação independentes; executa testes de desempenho, segurança e regressão e produz resultados verificáveis" },
        req: { en: ["Programming background in Python", "Familiar with independent evaluation sets and data isolation requirements; understands the separation of evaluation sets from training data and the prevention of contamination and leakage", "Skilled in baseline comparison, metric analysis and result reproduction; able to produce evaluation reports and locate performance differences", "Familiarity with common benchmarks and metrics (such as MMLU, GSM8K, HumanEval, C-Eval, BLEU, ROUGE and WER) preferred", "Familiarity with evaluation frameworks such as LM Evaluation Harness and OpenCompass preferred", "Model evaluation experience preferred"],
               pt: ["Base de programação em Python", "Familiaridade com conjuntos de avaliação independentes e requisitos de isolamento de dados; compreende a separação entre conjuntos de avaliação e dados de treinamento e a prevenção de contaminação e vazamento", "Domínio de comparação com linhas de base, análise de métricas e reprodução de resultados; capaz de produzir relatórios de avaliação e localizar diferenças de desempenho", "Familiaridade com benchmarks e métricas comuns (como MMLU, GSM8K, HumanEval, C-Eval, BLEU, ROUGE e WER) é um diferencial", "Familiaridade com frameworks de avaliação como LM Evaluation Harness e OpenCompass é um diferencial", "Experiência em avaliação de modelos é um diferencial"] },
        col: { en: ["Counterpart of iFLYTEK's model evaluation staff; confirms test sets, parameters, metrics and how results are recorded", "Performs RNP-side evaluation and independent re-checks; submits divergent samples and supports regression validation"],
               pt: ["Contraparte da equipe de avaliação de modelos da iFLYTEK; confirma conjuntos de teste, parâmetros, métricas e a forma de registro dos resultados", "Executa a avaliação do lado da RNP e a reverificação independente; submete amostras divergentes e apoia a validação de regressão"] } },
      { name: { en: "Platform, API and Application Test Engineer", pt: "Engenheiro de Testes de Plataforma, API e Aplicação" }, en: "Platform, API and Application Test Engineer", c: 2, e: 1,
        one: { en: "Verifies platform, API and multi-device application functions; tracks defects and provides acceptance evidence", pt: "Verifica as funcionalidades da plataforma, das APIs e da aplicação multiterminal; acompanha defeitos e fornece evidências para o aceite" },
        req: { en: ["Experience in software testing and test case design", "Familiar with API, authentication and permission testing", "Able to run multi-device compatibility testing, defect tracking and regression validation"],
               pt: ["Experiência em testes de software e desenho de casos de teste", "Familiaridade com testes de API, autenticação e permissões", "Capaz de conduzir testes de compatibilidade multiterminal, rastreamento de defeitos e validação de regressão"] },
        col: { en: ["Counterpart of iFLYTEK's test and implementation teams; confirms the acceptance test cases for the product, APIs and application", "Performs RNP-side testing; submits reproducible defects and confirms the fixes"],
               pt: ["Contraparte das equipes de testes e implantação da iFLYTEK; confirma os casos de teste de aceite do produto, das APIs e da aplicação", "Executa os testes do lado da RNP; submete defeitos reproduzíveis e confirma as correções"] } },
    ] },
  { name: { en: "Application and Operations", pt: "Aplicação e Operações" },
    intro: { en: "Responsible for integrating the National AI Application with RNP systems, for user approval, content and day-to-day operations, and for maintaining the delivered code that RNP is authorised to modify. Work starts in project phase 3.",
             pt: "Responsável pela integração da Aplicação Nacional de IA com os sistemas da RNP, pela aprovação de usuários, pelo conteúdo e pela operação diária, e pela manutenção do código entregue que a RNP está autorizada a modificar. O trabalho começa na fase 3 do projeto." },
    roles: [
      { name: { en: "Software Developer", pt: "Desenvolvedor de Software" }, en: "Software Developer", c: 2, e: 2,
        one: { en: "Integrates the National AI Application with RNP systems; develops local content and features", pt: "Integra a Aplicação Nacional de IA aos sistemas da RNP; desenvolve conteúdos e funcionalidades locais" },
        req: { en: ["Mobile (Android, iOS) and web front-end and back-end development skills", "Experience in system integration and API development", "Able to integrate against documented APIs"],
               pt: ["Habilidades em desenvolvimento mobile (Android, iOS) e web, front-end e back-end", "Experiência em integração de sistemas e desenvolvimento de APIs", "Capaz de realizar integrações com base em APIs documentadas"] },
        col: { en: ["Counterpart of iFLYTEK's application team: connects RNP systems according to the confirmed interface list; local feature development and the use of source code follow the scope agreed by both parties", "After shadowing, takes over RNP-side integration, and development and releases within the authorised scope"],
               pt: ["Contraparte da equipe de aplicação da iFLYTEK: integra os sistemas da RNP conforme a lista de interfaces confirmada; o desenvolvimento de funcionalidades locais e o uso do código-fonte seguem o escopo acordado entre as partes", "Após o acompanhamento, assume a integração do lado da RNP e o desenvolvimento e a publicação dentro do escopo autorizado"] } },
      { name: { en: "Application Operations Engineer", pt: "Engenheiro de Operações da Aplicação" }, en: "Application Operations Engineer", c: 2, e: 3,
        one: { en: "Operates the application: configuration, user approval, API credentials and usage analytics", pt: "Opera a aplicação: configuração, aprovação de usuários, credenciais de API e análise de uso" },
        req: { en: ["Experience operating internet products or government digital services", "Able to handle user approval and permission management", "Able to read usage data and report issues"],
               pt: ["Experiência na operação de produtos digitais ou serviços públicos digitais", "Capaz de tratar a aprovação de usuários e a gestão de permissões", "Capaz de interpretar dados de uso e reportar problemas"] },
        col: { en: ["Counterpart of iFLYTEK's application and support teams: operates in parallel during the assisted O&M period, then independently"],
               pt: ["Contraparte das equipes de aplicação e suporte da iFLYTEK: opera em paralelo durante o período de operação assistida e, depois, de forma independente"] } },
      { name: { en: "Solution Architect", pt: "Arquiteto de Soluções" }, en: "Solution Architect", c: 1, e: 1,
        one: { en: "Owns the application architecture and the integration design with RNP systems", pt: "Responsável pela arquitetura da aplicação e pelo desenho de integração com os sistemas da RNP" },
        req: { en: ["Enterprise system architecture experience; familiar with cloud, containers, microservices and API gateways", "Understands model service invocation and security requirements"],
               pt: ["Experiência em arquitetura de sistemas corporativos; familiaridade com nuvem, contêineres, microsserviços e gateways de API", "Compreende a invocação de serviços de modelos e os requisitos de segurança"] },
        col: { en: ["Counterpart of iFLYTEK's architects: jointly defines the application architecture, integration points and security design; reviews interface and deployment designs"],
               pt: ["Contraparte dos arquitetos da iFLYTEK: define em conjunto a arquitetura da aplicação, os pontos de integração e o desenho de segurança; revisa os desenhos de interfaces e de implantação"] } },
    ] },
  { name: { en: "Platform O&M", pt: "Operação e Manutenção da Plataforma" },
    intro: { en: "Runs the National AI Capability Center platform after acceptance. Work starts in project phase 3; all positions belong to the core team. Data-centre-level operations staff on the Huawei side are outside the scope of this reference.",
             pt: "Opera a plataforma do Centro Nacional de Capacidade em IA após o aceite. O trabalho começa na fase 3 do projeto; todas as posições pertencem à equipe núcleo. O pessoal de operação no nível do data center, do lado da Huawei, está fora do escopo desta referência." },
    roles: [
      { name: { en: "O&M Lead", pt: "Líder de Operações" }, en: "O&M Lead", c: 1, e: 0,
        one: { en: "Coordinates O&M processes, staff assignments, incident escalation and changes to keep services running in an orderly way", pt: "Coordena os processos de operação, a distribuição de tarefas, o escalonamento de incidentes e as mudanças, garantindo a operação ordenada dos serviços" },
        req: { en: ["Experience in platform operations and IT service management", "Familiar with incident, change and escalation processes", "Understands capacity, backup and recovery, and continuity management"],
               pt: ["Experiência em operação de plataformas e gestão de serviços de TI", "Familiaridade com processos de incidentes, mudanças e escalonamento", "Compreende gestão de capacidade, backup e recuperação e continuidade"] },
        col: { en: ["Counterpart of iFLYTEK's O&M and support teams; confirms the division of service responsibilities, escalation channels and change mechanisms", "Organises RNP's O&M staff to take over operations; reviews operation reports, open items and handover results"],
               pt: ["Contraparte das equipes de operação e suporte da iFLYTEK; confirma a divisão de responsabilidades de serviço, os canais de escalonamento e os mecanismos de mudança", "Organiza a equipe de operação da RNP para assumir as operações; revisa relatórios de operação, pendências e resultados da transferência"] } },
      { name: { en: "Platform and Application O&M Engineer", pt: "Engenheiro de Operações de Plataforma e Aplicação" }, en: "Platform and Application O&M Engineer", c: 3, e: 0,
        one: { en: "Performs inspections, alert handling, configuration backups and approved maintenance operations for the platform, application and model services", pt: "Executa inspeções, tratamento de alertas, backups de configuração e operações de manutenção aprovadas para a plataforma, a aplicação e os serviços de modelos" },
        req: { en: ["Familiar with Linux, containers and application service maintenance", "Able to analyse logs, monitoring data and common alerts", "Skilled in configuration backup, recovery and patch changes"],
               pt: ["Familiaridade com Linux, contêineres e manutenção de serviços de aplicação", "Capaz de analisar logs, monitoramento e alertas comuns", "Domínio de backup e recuperação de configurações e aplicação de patches"] },
        col: { en: ["Counterpart of iFLYTEK's O&M team; takes part in inspections, configuration checks, recovery drills and approved maintenance operations", "Carries out day-to-day O&M of the platform, application and model services on RNP's side; performs releases, configuration changes and rollbacks as approved; submits fault details and supports remote diagnosis and verification of fixes"],
               pt: ["Contraparte da equipe de operação da iFLYTEK; participa de inspeções, verificações de configuração, exercícios de recuperação e operações de manutenção aprovadas", "Realiza a operação diária da plataforma, da aplicação e dos serviços de modelos do lado da RNP; executa publicações, mudanças de configuração e reversões conforme aprovado; submete informações de falhas e apoia o diagnóstico remoto e a verificação das correções"] } },
      { name: { en: "Security Operations Engineer", pt: "Engenheiro de Operações de Segurança" }, en: "Security Operations Engineer", c: 2, e: 0,
        one: { en: "Operates the security products; analyses alerts and tracks vulnerability remediation and product-side security incidents", pt: "Opera os produtos de segurança; analisa alertas e acompanha a correção de vulnerabilidades e os incidentes de segurança do lado do produto" },
        req: { en: ["Familiar with WAF, bastion host and audit security products", "Log analysis, vulnerability scanning and remediation tracking skills", "Experience in security incident investigation and coordinated response"],
               pt: ["Familiaridade com WAF, bastion host e produtos de segurança de auditoria", "Habilidades em análise de logs, varredura de vulnerabilidades e acompanhamento de correções", "Experiência em investigação de incidentes de segurança e resposta coordenada"] },
        col: { en: ["Counterpart of iFLYTEK's security product and support teams; takes part in deployment configuration, log integration and functional verification", "Carries out day-to-day security operations on RNP's side; provides incident details and works with iFLYTEK on product issues"],
               pt: ["Contraparte das equipes de produtos de segurança e suporte da iFLYTEK; participa da configuração de implantação, integração de logs e verificação funcional", "Realiza as operações diárias de segurança do lado da RNP; fornece informações de incidentes e atua em conjunto com a iFLYTEK nos problemas dos produtos"] } },
    ] },
];
GROUPS.forEach(g => { g.c = g.roles.reduce((a, r) => a + r.c, 0); g.e = g.roles.reduce((a, r) => a + r.e, 0); g.n = g.c + g.e; });
const CORE = GROUPS.reduce((a, g) => a + g.c, 0), EXT = GROUPS.reduce((a, g) => a + g.e, 0), TOTAL = CORE + EXT;
const NROLES = GROUPS.reduce((a, g) => a + g.roles.length, 0);
const fill = s => L(s).replace("{T}", TOTAL).replace("{G}", GROUPS.length).replace("{R}", NROLES).replace("{C}", CORE).replace("{E}", EXT);

// =====================================================================================
// Builders
// =====================================================================================
const R = (t, o = {}) => new TextRun({ text: t, font: F, size: o.size || 21, bold: !!o.bold, color: o.color || BLACK, italics: !!o.italics });
function runs(t, o = {}) { return String(t).split("**").map((seg, i) => seg ? R(seg, { ...o, bold: o.bold || i % 2 === 1 }) : null).filter(Boolean); }
const P = (t, o = {}) => D.push(new Paragraph({ children: runs(t, o), spacing: { after: o.after == null ? 120 : o.after, line: 310 } }));
const H1 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 420, after: 160 }, keepNext: true, children: [R(t, { size: 28, bold: true, color: NAVY })] }));
const split = (c, e) => `${L(UI.core)} ${c} · ${L(UI.ext)} ${e}`;
function H2(num, name, c, e, intro) {
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 40 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }],
    children: [R(`${num}  ${name}`, { size: 25, bold: true, color: NAVY }), R(`\t${split(c, e)}`, { size: 20, bold: true, color: BLUE })] }));
  D.push(new Paragraph({ spacing: { after: 140 }, keepNext: true, children: [R(intro, { size: 19, color: GRAY })] }));
}
function H3(num, name, en, c, e, note) {
  const kids = [R(`${num}  ${name}`, { size: 22, bold: true, color: BLACK })];
  const gloss = LANG === "pt" && en && en !== name ? en : null;
  const glossBelow = gloss && (name.length + gloss.length) > 75;   // long PT names: English name on its own line so the heading never wraps
  if (gloss && !glossBelow) kids.push(R(`   ${gloss}`, { size: 18, color: GRAY }));
  kids.push(R(`\t${split(c, e)}`, { size: 20, bold: true, color: BLUE }));
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: (note || glossBelow) ? 20 : 60 }, keepNext: true, tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }], children: kids }));
  if (glossBelow) D.push(new Paragraph({ spacing: { after: note ? 20 : 60 }, keepNext: true, children: [R(gloss, { size: 18, color: GRAY })] }));
  if (note) D.push(new Paragraph({ spacing: { after: 80 }, keepNext: true, children: [R(note, { size: 18, color: GRAY, italics: true })] }));
}
const TH = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function cell(children, w, o = {}) {
  return new TableCell({ children, width: { size: w, type: WidthType.DXA }, borders: { top: TH, bottom: TH, left: NONE, right: NONE },
    margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.CENTER,
    ...(o.rowSpan ? { rowSpan: o.rowSpan } : {}), ...(o.colSpan ? { columnSpan: o.colSpan } : {}),
    ...(o.fill ? { shading: { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } } : {}) });
}
const tp = (t, o = {}) => new Paragraph({ children: runs(t, { size: o.size || 19, bold: o.bold, color: o.color }), spacing: { after: 20, line: 260 }, alignment: o.align || AlignmentType.LEFT });
const bp = t => new Paragraph({ children: runs(t, { size: 19 }), numbering: { reference: "cellb", level: 0 }, spacing: { after: 30, line: 270 } });
const C = AlignmentType.CENTER;

// ---- cover ----
D.push(new Paragraph({ spacing: { after: 100 }, children: [R(L(UI.title), { size: 34, bold: true, color: RED })] }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [R(L(UI.subtitle), { size: 26, bold: true, color: NAVY })] }));
D.push(new Paragraph({ spacing: { after: 320 }, children: [R(L(UI.meta), { size: 18, color: GRAY })] }));
P(L(UI.intro1)); P(L(UI.intro2), { after: 60 });

// ---- 1 ----
H1(L(UI.s1));
P(fill(UI.s1p0)); P(fill(UI.s1core)); P(fill(UI.s1ext)); P(fill(UI.s1last), { after: 60 });

// ---- 2 ----
H1(L(UI.s2));
const W = LANG === "pt" ? [600, 1700, 2150, 800, 800, 3588] : [600, 1700, 2150, 620, 640, 3928];
const hdr = L(UI.hdr);
const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: hdr.map((h, i) => cell([tp(h, { bold: true, size: 19, color: NAVY, align: [0, 3, 4].includes(i) ? C : AlignmentType.LEFT })], W[i], { fill: "DCE6F1" })) })];
GROUPS.forEach((g, gi) => {
  g.roles.forEach((r, ri) => {
    const c = [];
    c.push(cell([tp(`${gi + 1}.${ri + 1}`, { align: C, color: GRAY })], W[0]));
    if (ri === 0) c.push(cell([tp(`**${gi + 1}  ${L(g.name)}**`, { color: NAVY }), tp(split(g.c, g.e), { size: 17, color: BLUE })], W[1], { rowSpan: g.roles.length, fill: LIGHT }));
    c.push(cell([tp(`**${L(r.name)}**`)], W[2]));
    c.push(cell([tp(r.c ? `**${r.c}**` : "–", { align: C, color: r.c ? NAVY : GRAY })], W[3]));
    c.push(cell([tp(r.e ? `${r.e}` : "–", { align: C, color: r.e ? BLUE : GRAY })], W[4]));
    c.push(cell([tp(L(r.one), { size: 18 })], W[5]));
    rows.push(new TableRow({ cantSplit: true, children: c }));
  });
});
rows.push(new TableRow({ cantSplit: true, children: [
  cell([tp(`**${L(UI.total)}**`, { color: NAVY })], W[0] + W[1] + W[2], { colSpan: 3, fill: LIGHT }),
  cell([tp(`**${CORE}**`, { align: C, color: NAVY, size: 21 })], W[3], { fill: LIGHT }),
  cell([tp(`**${EXT}**`, { align: C, color: BLUE, size: 21 })], W[4], { fill: LIGHT }),
  cell([tp(`**${fill(UI.totalLine)}**`, { color: RED })], W[5], { fill: LIGHT }),
] }));
D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: W, rows }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

// ---- 3 ----
H1(L(UI.s3));
P(L(UI.s3a), { after: 20 }); P(L(UI.s3b), { after: 20 }); P(L(UI.s3c), { after: 60 });
const WR = [1500, PAGE_W - 1500];
GROUPS.forEach((g, gi) => {
  H2(`${gi + 1}`, L(g.name), g.c, g.e, L(g.intro));
  g.roles.forEach((r, ri) => {
    H3(`${gi + 1}.${ri + 1}`, L(r.name), r.en, r.c, r.e, r.note ? L(r.note) : null);
    D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: WR, rows: [
      new TableRow({ cantSplit: true, children: [ cell([tp(L(UI.req), { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }), cell(L(r.req).map(bp), WR[1], { valign: VerticalAlign.TOP }) ] }),
      new TableRow({ cantSplit: true, children: [ cell([tp(L(UI.col), { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }), cell(L(r.col).map(bp), WR[1], { valign: VerticalAlign.TOP }) ] }),
    ] }));
    D.push(new Paragraph({ spacing: { after: 10 }, children: [] }));
  });
});

// =====================================================================================
const bullet = (ref, left, hanging) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });
const doc = new Document({
  creator: "iFLYTEK", title: `${L(UI.title)}: ${L(UI.subtitle)}`,
  styles: { default: { document: { run: { font: F, size: 21, color: BLACK } } } },
  numbering: { config: [bullet("cellb", 260, 220)] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      R(L(UI.footer), { size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 16, color: GRAY }) ] })] }) },
    children: D,
  }],
});
const OUT = L(UI.out);
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes", "core", CORE, "ext", EXT, "total", TOTAL, "roles", NROLES); });
