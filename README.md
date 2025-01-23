Relatórios Automatizados

Este repositório contém um script Python desenvolvido para automatizar a geração de relatórios dinâmicos em planilhas Excel. Ele permite o processamento de dados de entrada, a criação de índices e tabelas organizadas, e o envio automatizado de e-mails com os arquivos gerados.

Funcionalidades

Processamento de Dados:

Identifica automaticamente o arquivo mais recente na pasta de entrada.

Realiza validações para garantir que os dados necessários estão presentes.

Gera tabelas dinâmicas organizadas por categoria e modelo de produto.

Geração de Relatórios:

Cria um arquivo Excel com múltiplas abas baseadas nos dados processados.

Inclui uma aba de índice com hiperlinks para facilitar a navegação.

Aplica estilos personalizados e formatação visual para melhorar a apresentação.

Automação de E-mail:

Integração com servidor SMTP para envio automático de relatórios por e-mail (lista de destinatários pode ser ajustada conforme necessidade).

Configuração

Requisitos

Python 3.8 ou superior

Bibliotecas Python:

pandas

openpyxl

smtplib

email

Instalação

Clone este repositório:

git clone https://github.com/seu-usuario/relatorios-automatizados.git

Navegue até o diretório do projeto:

cd relatorios-automatizados

Instale as dependências:

pip install pandas openpyxl

Configuração de Caminhos

No script, configure os caminhos das pastas de entrada e saída conforme sua estrutura local:

pasta_entrada: Diretório onde os arquivos de entrada serão armazenados.

caminho_saida: Local onde o arquivo Excel gerado será salvo.

Exemplo:

pasta_entrada = r"C:\Projeto\localizador\entrada"
caminho_saida = r"C:\Projeto\localizador\saida\LOCALIZADOR_GM.xlsx"

Configuração do Servidor de E-mail

Substitua as credenciais e configurações do servidor SMTP para sua conta de e-mail:

SMTP_SERVER = 'smtp.seuservidor.com'
SMTP_PORT = 587
EMAIL_USER = 'seu.email@seuservidor.com'
EMAIL_PASS = 'sua_senha'

Como Usar

Certifique-se de que os arquivos de entrada estão na pasta configurada.

Execute o script:

python script.py

O relatório será gerado no caminho configurado e poderá ser enviado automaticamente por e-mail (se configurado).

Estrutura do Relatório

Índice:

Hiperlinks para acessar abas específicas.

Categorias definidas para facilitar a organização.

Abas de Produto:

Tabelas dinâmicas organizadas por linha de produto, cor e status do veículo.

Dados sumarizados por categoria.

Personalização

Mapeamento de Categorias: O dicionário categorias pode ser ajustado para incluir novas categorias ou modificar existentes.

Estilo Visual: O script utiliza formatações padrão que podem ser alteradas diretamente no código para refletir suas preferências visuais.

Contribuição

Contribuições são bem-vindas! Caso encontre problemas ou tenha sugestões de melhorias, abra uma issue ou envie um pull request.

Licença

Este projeto está licenciado sob a Licença MIT. 

Consulte o arquivo LICENSE para mais informações.
