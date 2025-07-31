Sistema de Gerenciamento de Metas em VB6 com Integração a Banco de Dados SQL Server.

Desenvolvido em Visual Basic 6 (VB6), este sistema permite o controle e a gestão completa de metas. Ele se integra a um banco de dados SQL Server para persistência das informações, oferecendo funcionalidades robustas de CRUD (Criar, Ler, Atualizar, Excluir) para as metas.

Principais funcionalidades:

Gerenciamento Abrangente de Metas: Adição, edição, exclusão e visualização detalhada de metas individuais.

Detalhes da Meta: Cada meta pode incluir:

Descrição (limitada a 255 caracteres)

Prioridade (Baixa, Média, Alta)

Data Prevista de Vencimento

Status de Conclusão (indicando se a meta foi concluída ou não)

Data de Conclusão (exibida apenas para metas concluídas)

Cálculo de Dias de Atraso (se aplicável)

Interface Intuitiva com FP Spread: Utiliza o componente FarPoint Spread (FP Spread) para uma visualização clara e organizada das metas em formato de grade.

Navegação e Edição Controladas: O grid permite a visualização detalhada das metas, e possui um botão "Consultar" por linha que carrega os dados da meta selecionada para os campos de edição no formulário, impedindo a edição direta no grid para manter a integridade dos dados e facilitar o fluxo de trabalho.

Experiência do Usuário Aprimorada: Inclui melhorias visuais e técnicas nos botões e na visualização geral da tela para uma interação mais fluida e eficiente.

Arquitetura Refatorada: O código principal do sistema foi refatorado utilizando módulos e classes para uma maior organização, manutenibilidade e escalabilidade.

Este sistema oferece uma solução eficiente para o acompanhamento e a gestão de metas, combinando uma interface funcional com uma estrutura de código aprimorada.
