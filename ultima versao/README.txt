SISTEMA ESCOLAR — VERSÃO AMPLIADA

COMO ABRIR NO VS CODE
1. Extraia o arquivo ZIP para uma pasta.
2. Abra a pasta no Visual Studio Code (Arquivo > Abrir Pasta).
3. Abra o terminal integrado.
4. (Opcional, recomendado) Crie ambiente virtual: python -m venv .venv
   Windows: .venv\Scripts\activate
   Linux/macOS: source .venv/bin/activate
5. Instale as dependências: pip install -r requirements.txt
6. Execute: python app.py
7. Acesse: http://127.0.0.1:5000

PRIMEIRO ACESSO
A primeira conta cadastrada recebe o perfil Administrador. As contas seguintes recebem perfil Professor.

RECURSOS
- Login de professores e perfis de acesso (administrador/professor).
- Cadastro, listagem, notas, frequência e ranking de alunos.
- Dashboard com indicadores e alertas.
- Calendário escolar com cadastro e exclusão de eventos.
- Registro do histórico de comunicação com responsáveis (não envia mensagens automaticamente).
- Boletim individual em PDF.
- Exportação de relatórios em CSV e Excel.
- Auditoria de atividades para administrador.
- Backup manual do banco SQLite para administrador.
- Interface responsiva para computador, tablet e celular.
- Alternância de tema claro/escuro (preferência salva no navegador).

IMPORTANTE
- A primeira conta deve ser criada pela pessoa responsável pela administração.
- Antes de publicar na internet, altere SECRET_KEY por uma chave longa e aleatória e configure HTTPS.
- O backup é manual: baixe periodicamente o arquivo pelo menu de administrador e guarde-o em local seguro.
- O sistema é uma base funcional de projeto. Ainda não inclui recuperação de senha por e-mail, envio automático de mensagens, exclusão lógica/recuperação de registros nem agendamento automático de backups. Para uso real em escola, recomenda-se revisão de segurança, política de privacidade, controle de acesso por turma e adequação completa à LGPD.
- O calendário e os contatos são registrados por professor logado.
