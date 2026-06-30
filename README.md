const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
        ShadingType, PageBreak } = require('docx');
const fs = require('fs');

const PURPLE = '6C3EFF';
const YELLOW = 'FFE600';
const BLACK  = '0A0A0A';
const GRAY   = '5E5C56';
const GREEN  = '00B87A';
const WHITE  = 'FFFFFF';
const LIGHT_PURPLE = 'EDE8FF';
const LIGHT_GREEN  = 'D6F7ED';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'E0DED7' };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(texto) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: PURPLE, space: 8 } },
    children: [new TextRun({ text: texto, bold: true, size: 36, font: 'Arial', color: BLACK })]
  });
}

function h2(texto) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text: texto, bold: true, size: 28, font: 'Arial', color: PURPLE })]
  });
}

function h3(texto) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: texto, bold: true, size: 24, font: 'Arial', color: BLACK })]
  });
}

function p(texto, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text: texto, size: 22, font: 'Arial', color: GRAY, ...options })]
  });
}

function bullet(texto, nivel = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: nivel },
    spacing: { before: 40, after: 60 },
    children: [new TextRun({ text: texto, size: 22, font: 'Arial', color: GRAY })]
  });
}

function numbered(texto) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text: texto, size: 22, font: 'Arial', color: BLACK })]
  });
}

function destaque(titulo, corpo, cor = LIGHT_PURPLE, corTexto = PURPLE) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: cor, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: titulo, bold: true, size: 22, font: 'Arial', color: corTexto })] }),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: corpo, size: 20, font: 'Arial', color: BLACK })] })
        ]
      })]
    })]
  });
}

function tabelaTecnologias() {
  const linhas = [
    ['Camada', 'Tecnologia', 'Arquivo principal'],
    ['Frontend', 'HTML5 + CSS3 + JavaScript', 'frontend/index.html'],
    ['Backend', 'PHP 8.x', 'backend/auth.php, missoes.php, mentor.php'],
    ['Banco de Dados', 'MySQL 8.x', 'sql/ia_boost.sql'],
    ['API IA', 'Anthropic Claude API', 'backend/mentor.php'],
    ['Hospedagem', 'InfinityFree / Hostinger', 'cPanel + phpMyAdmin'],
    ['Versionamento', 'Git + GitHub', '.gitignore + README.md'],
  ];

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2800, 4560],
    rows: linhas.map((linha, i) => new TableRow({
      children: linha.map(cel => new TableCell({
        borders,
        width: { size: 0, type: WidthType.AUTO },
        shading: { fill: i === 0 ? PURPLE : (i % 2 === 0 ? 'F8F8F7' : WHITE), type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: cel, size: 20, font: 'Arial', bold: i === 0, color: i === 0 ? WHITE : BLACK })] })]
      }))
    }))
  });
}

function quebra() {
  return new Paragraph({ children: [new PageBreak()] });
}

function espaco() {
  return new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] });
}

const doc = new Document({
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }, { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
      { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22, color: GRAY } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', run: { size: 36, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', run: { size: 28, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } }
    },
    children: [
      // CAPA
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1440, after: 200 }, children: [new TextRun({ text: '⚡', size: 80, font: 'Arial' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 }, children: [new TextRun({ text: 'IA Boost', bold: true, size: 64, font: 'Arial', color: PURPLE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Aprenda IA. Ganhe tempo. Aumente sua renda.', size: 28, font: 'Arial', color: GRAY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 400 }, children: [new TextRun({ text: '"O Duolingo da Inteligência Artificial"', size: 24, font: 'Arial', color: GRAY, italics: true })] }),
      espaco(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [new TableRow({
          children: [
            new TableCell({ borders: noBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: LIGHT_PURPLE, type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 200, right: 200 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '🎯 Projeto Final', bold: true, size: 22, font: 'Arial', color: PURPLE })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Desenvolvimento de MVP', size: 20, font: 'Arial', color: BLACK })] })] }),
            new TableCell({ borders: noBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: 'FFF9CC', type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 200, right: 200 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '💻 Stack', bold: true, size: 22, font: 'Arial', color: '8B7500' })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'HTML · PHP · MySQL', size: 20, font: 'Arial', color: BLACK })] })] }),
            new TableCell({ borders: noBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 200, right: 200 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '🚀 Diferencial', bold: true, size: 22, font: 'Arial', color: '006B47' })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Mentor com IA real', size: 20, font: 'Arial', color: BLACK })] })] }),
          ]
        })]
      }),
      espaco(),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 0 }, children: [new TextRun({ text: '2025', size: 22, font: 'Arial', color: GRAY })] }),

      quebra(),

      // 1. VISÃO GERAL
      h1('1. Visão Geral do Projeto'),
      p('A IA Boost é uma plataforma de aprendizado prático de Inteligência Artificial. Diferente de plataformas tradicionais que ensinam teoria, a IA Boost foca em resultados concretos: o usuário aprende fazendo, completando missões práticas que geram valor real em sua vida profissional.'),
      espaco(),
      h2('Problema identificado'),
      p('A maioria das pessoas sabe que deveria usar IA no trabalho, mas não sabe por onde começar. Plataformas existentes são genéricas demais ou técnicas demais para o público geral.'),
      espaco(),
      h2('Público-alvo'),
      bullet('Profissionais autônomos (lash designers, corretores, professores, contadores)'),
      bullet('Pequenos empreendedores'),
      bullet('Estudantes e recém-formados'),
      bullet('Pessoas que querem renda extra usando IA'),
      espaco(),
      h2('Solução proposta'),
      p('Uma plataforma gamificada no estilo Duolingo onde o usuário:'),
      bullet('Escolhe seu objetivo (ganhar dinheiro, criar conteúdo, vender mais...)'),
      bullet('Recebe uma jornada de missões personalizada'),
      bullet('Completa missões práticas com prompts prontos para usar'),
      bullet('Conversa com um Mentor IA que responde dúvidas específicas de sua profissão'),
      bullet('Ganha XP, sobe de nível e acompanha seu progresso'),

      quebra(),

      // 2. ARQUITETURA
      h1('2. Arquitetura e Tecnologias'),
      p('O sistema segue a arquitetura MVC (Model-View-Controller) com separação clara entre frontend, backend e banco de dados.'),
      espaco(),
      tabelaTecnologias(),
      espaco(),
      h2('Estrutura de arquivos do projeto'),
      destaque(
        '📁 ia-boost/ (raiz do projeto)',
        'frontend/\n   index.html        ← Interface principal (HTML + CSS + JS)\nbackend/\n   config.php        ← Conexão com banco + funções utilitárias\n   auth.php          ← API de login e registro\n   missoes.php       ← API de missões e progresso\n   mentor.php        ← API do Mentor IA (integração Claude)\nsql/\n   ia_boost.sql      ← Schema do banco de dados + dados iniciais\n.gitignore           ← Arquivos ignorados pelo Git\nREADME.md            ← Documentação do projeto',
        'F8F8F7', BLACK
      ),

      quebra(),

      // 3. BANCO DE DADOS
      h1('3. Banco de Dados'),
      h2('Diagrama das tabelas'),
      p('O banco de dados possui 4 tabelas principais:'),
      espaco(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'Tabela', bold: true, size: 20, font: 'Arial', color: WHITE })] })] }), new TableCell({ borders, width: { size: 7160, type: WidthType.DXA }, shading: { fill: PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'Descrição e campos principais', bold: true, size: 20, font: 'Arial', color: WHITE })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'usuarios', bold: true, size: 20, font: 'Arial', color: PURPLE })] })] }), new TableCell({ borders, shading: { fill: 'FAFAF8', type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'id, nome, email, senha (hash), objetivo, xp, sequencia_dias, criado_em', size: 20, font: 'Arial', color: BLACK })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'missoes', bold: true, size: 20, font: 'Arial', color: PURPLE })] })] }), new TableCell({ borders, shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'id, titulo, descricao, objetivo, xp_recompensa, tempo_estimado, ordem', size: 20, font: 'Arial', color: BLACK })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'progresso_missoes', bold: true, size: 20, font: 'Arial', color: PURPLE })] })] }), new TableCell({ borders, shading: { fill: 'FAFAF8', type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'id, usuario_id (FK), missao_id (FK), status, resultado, concluida_em', size: 20, font: 'Arial', color: BLACK })] })] })] }),
          new TableRow({ children: [new TableCell({ borders, shading: { fill: LIGHT_PURPLE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'mentor_historico', bold: true, size: 20, font: 'Arial', color: PURPLE })] })] }), new TableCell({ borders, shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'id, usuario_id (FK), papel (user/assistant), mensagem, criado_em', size: 20, font: 'Arial', color: BLACK })] })] })] }),
        ]
      }),

      quebra(),

      // 4. COMO SUBIR NA HOSPEDAGEM
      h1('4. Como Subir na Hospedagem (Passo a Passo)'),
      destaque('💡 Hospedagem gratuita recomendada', 'InfinityFree (infinityfree.com) — PHP + MySQL grátis, sem limite de tráfego. Alternativa: 000webhost.com', LIGHT_GREEN, '006B47'),
      espaco(),
      h2('Passo 1 — Criar conta na hospedagem'),
      numbered('Acesse infinityfree.com e crie uma conta gratuita'),
      numbered('Clique em "Create Account" e escolha um subdomínio (ex: ia-boost.rf.gd)'),
      numbered('Anote: host do banco, usuário MySQL, senha MySQL'),
      espaco(),
      h2('Passo 2 — Configurar o banco de dados'),
      numbered('No painel da hospedagem, abra o phpMyAdmin'),
      numbered('Clique em "Novo banco de dados" → nomeie ia_boost → clique Criar'),
      numbered('Clique no banco criado → aba "SQL" → cole todo o conteúdo do arquivo ia_boost.sql'),
      numbered('Clique em "Executar" — as tabelas serão criadas automaticamente'),
      espaco(),
      h2('Passo 3 — Configurar o backend'),
      p('Abra o arquivo backend/config.php e edite as linhas:'),
      destaque(
        '⚙️ config.php — altere estes valores',
        "define('DB_HOST', 'sql.seuservidor.com');  // Host do painel\ndefine('DB_NAME', 'ia_boost');\ndefine('DB_USER', 'epiz_XXXXXXX');  // Usuário do painel\ndefine('DB_PASS', 'SUA_SENHA');",
        'FFF9CC', '8B7500'
      ),
      espaco(),
      h2('Passo 4 — Configurar a chave da API de IA'),
      numbered('Acesse console.anthropic.com e crie uma conta'),
      numbered('Vá em "API Keys" → crie uma chave nova'),
      numbered('Abra backend/mentor.php e substitua SUA_CHAVE_ANTHROPIC_AQUI pela chave gerada'),
      espaco(),
      h2('Passo 5 — Fazer upload dos arquivos'),
      numbered('No painel da hospedagem, acesse o "File Manager" (gerenciador de arquivos)'),
      numbered('Abra a pasta htdocs (ou public_html)'),
      numbered('Crie uma pasta chamada ia-boost'),
      numbered('Faça upload de todos os arquivos dentro da pasta ia-boost/'),
      numbered('Abra frontend/index.html e troque a URL do API_BASE pelo seu domínio'),
      espaco(),
      h2('Passo 6 — Testar o sistema'),
      numbered('Acesse https://SEU_DOMINIO/ia-boost/frontend/index.html'),
      numbered('Crie uma conta de teste'),
      numbered('Complete uma missão e verifique se o XP é salvo'),
      numbered('Teste o Mentor IA enviando uma mensagem'),

      quebra(),

      // 5. GIT E GITHUB
      h1('5. Versionamento com Git e GitHub'),
      p('O Git registra todo o histórico de alterações do projeto. O GitHub armazena o código online e permite trabalho em equipe.'),
      espaco(),
      h2('Comandos essenciais'),
      destaque(
        '💻 Terminal — execute na pasta raiz do projeto (ia-boost/)',
        '# Inicializar repositório\ngit init\ngit add .\ngit commit -m "feat: MVP inicial do IA Boost"\n\n# Conectar ao GitHub\ngit remote add origin https://github.com/SEU_USUARIO/ia-boost.git\ngit push -u origin main\n\n# A cada atualização\ngit add .\ngit commit -m "fix: descrição do que foi alterado"\ngit push',
        'F0EFEB', BLACK
      ),
      espaco(),
      h2('Arquivo .gitignore (crie na raiz)'),
      destaque(
        '📄 .gitignore — proteja dados sensíveis',
        '# Nunca envie senhas ao GitHub!\nbackend/config.php\n*.env\n.DS_Store\nnode_modules/',
        'FCEAEA', 'A32D2D'
      ),

      quebra(),

      // 6. PITCH
      h1('6. Pitch de Apresentação'),
      p('Use esta estrutura para apresentar o projeto. Tempo sugerido: 8 a 10 minutos.'),
      espaco(),

      h2('Slide 1 — Abertura (30 seg)'),
      destaque('🎯 Frase de impacto para abrir', '"E se você pudesse aprender a usar IA do mesmo jeito que aprendeu a usar o WhatsApp — na prática, errando, tentando — mas com alguém te guiando passo a passo?" Aguardem um segundo de silêncio após falar.', LIGHT_PURPLE, PURPLE),
      espaco(),

      h2('Slide 2 — Problema (1 min)'),
      bullet('83% dos brasileiros nunca usaram IA no trabalho (fonte: pesquisa FGV 2024)'),
      bullet('Os que tentam desistem porque os cursos são teóricos e genéricos'),
      bullet('Resultado: quem não aprende IA está perdendo oportunidades de renda e produtividade'),
      espaco(),

      h2('Slide 3 — Público impactado (30 seg)'),
      p('Qualquer pessoa que trabalha: lash designers, corretores, professores, contadores, autônomos, pequenos empreendedores, freelancers e estudantes.'),
      espaco(),

      h2('Slide 4 — Solução: IA Boost (2 min)'),
      bullet('Plataforma web gamificada — missões práticas com resultado concreto em cada etapa'),
      bullet('Jornada personalizada por objetivo e profissão'),
      bullet('Prompts prontos: o usuário copia, usa na IA favorita e traz o resultado de volta'),
      bullet('Mentor IA: conversa personalizada sobre a profissão do usuário'),
      bullet('Gamificação: XP, sequência diária, ranking — mantém o usuário engajado'),
      espaco(),

      h2('Slide 5 — Demonstração prática (3 min)'),
      numbered('Mostrar a tela de onboarding: escolher objetivo'),
      numbered('Mostrar a home com missões e progresso'),
      numbered('Abrir uma missão, copiar o prompt, colar resultado e ganhar XP'),
      numbered('Acessar o Mentor IA e enviar uma pergunta'),
      numbered('Mostrar o banco de dados registrando tudo em tempo real'),
      espaco(),

      h2('Slide 6 — Tecnologias utilizadas (1 min)'),
      bullet('Frontend: HTML5 + CSS3 + JavaScript puro (sem framework) — acessível em qualquer dispositivo'),
      bullet('Backend: PHP 8 com PDO — autenticação segura com hash de senhas'),
      bullet('Banco de dados: MySQL com 4 tabelas e chaves estrangeiras'),
      bullet('API de IA: Anthropic Claude — respostas personalizadas em tempo real'),
      bullet('Hospedagem: InfinityFree — deploy real, funcionando na internet'),
      bullet('Versionamento: Git + GitHub com commits organizados'),
      espaco(),

      h2('Slide 7 — Possibilidades futuras (1 min)'),
      bullet('Aplicativo móvel para Android e iOS'),
      bullet('Marketplace de prompts — usuários vendem seus melhores prompts'),
      bullet('Plano empresarial para treinamento de equipes'),
      bullet('Certificações reconhecidas pelo mercado'),
      bullet('Programa de afiliados com comissão recorrente'),
      espaco(),

      h2('Slide 8 — Fechamento (30 seg)'),
      destaque(
        '🚀 Mensagem final',
        '"A IA não vai substituir as pessoas. Vai substituir as pessoas que não sabem usar IA. O IA Boost resolve isso — de forma prática, acessível e com resultado real. Obrigado!"',
        LIGHT_PURPLE, PURPLE
      ),

      quebra(),

      // 7. CHECKLIST
      h1('7. Checklist de Entrega'),
      p('Verifique todos os itens antes de entregar o projeto:'),
      espaco(),
      h2('Requisitos técnicos obrigatórios'),
      destaque('📋 Todos os itens abaixo devem estar funcionando', '', 'FFF9CC', '8B7500'),
      bullet('Interface web com HTML e CSS estilizado'),
      bullet('Interatividade com JavaScript (login, missões, chat, XP)'),
      bullet('Backend PHP com pelo menos 3 arquivos de API'),
      bullet('Banco MySQL com 4 tabelas e relacionamentos (FK)'),
      bullet('Organização de arquivos em pastas (frontend, backend, sql)'),
      bullet('Repositório Git com pelo menos 3 commits'),
      bullet('README.md explicando como rodar o projeto'),
      espaco(),
      h2('Diferenciais que impressionam'),
      bullet('Integração com API de IA real (Claude/OpenAI)'),
      bullet('Autenticação com hash de senha (password_hash)'),
      bullet('Deploy funcionando online (não só local)'),
      bullet('Validação de dados no frontend E no backend'),
      bullet('Histórico de chat salvo no banco de dados'),

      espaco(),
      espaco(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 0 },
        children: [
          new TextRun({ text: '⚡ IA Boost — Projeto Final  ·  2025', size: 20, font: 'Arial', color: GRAY, italics: true })
        ]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/ia-boost/GUIA_COMPLETO_IA_BOOST.docx', buf);
  console.log('Documento criado com sucesso!');
});
