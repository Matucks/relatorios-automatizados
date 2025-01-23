# Relatórios Automatizados

Este repositório contém um script em Python criado para automatizar a geração de relatórios dinâmicos em planilhas Excel. O script processa dados de entrada, cria índices e tabelas organizadas e permite o envio automático de e-mails com os relatórios gerados.

---

## Funcionalidades

### 1. **Processamento de Dados**
- Identifica automaticamente o arquivo mais recente na pasta de entrada.
- Realiza validações para garantir a presença dos dados necessários.
- Gera tabelas dinâmicas organizadas por categoria e modelo de produto.

### 2. **Geração de Relatórios**
- Cria um arquivo Excel com múltiplas abas baseadas nos dados processados.
- Inclui uma aba de índice com hiperlinks para navegação fácil entre as abas.
- Aplica estilos personalizados e formatações visuais para melhorar a apresentação.

### 3. **Automatização de E-mail**
- Integra-se a um servidor SMTP para envio automático de relatórios por e-mail.
- Lista de destinatários configurável conforme necessidade.

---

## Configuração

### 1. **Requisitos**
- **Python 3.8 ou superior.**

### 2. **Bibliotecas Necessárias**
- `pandas`
- `openpyxl`
- `smtplib`
- `email`

Instale as dependências executando o comando abaixo:
```bash
pip install pandas openpyxl
```

### 3. **Configuração de Caminhos**

Configure os caminhos de entrada e saída no script conforme a estrutura local:
- **Pasta de Entrada**: Diretório onde os arquivos de entrada serão armazenados.
- **Caminho de Saída**: Local onde o arquivo Excel gerado será salvo.

Exemplo:
```python
pasta_entrada = r"C:\Projeto\localizador\entrada"
caminho_saida = r"C:\Projeto\localizador\saida\LOCALIZADOR_GM.xlsx"
```

### 4. **Configuração do Servidor de E-mail**

Substitua as credenciais e as configurações do servidor SMTP no script para sua conta de e-mail:
```python
SMTP_SERVER = 'smtp.seuservidor.com'
SMTP_PORT = 587
EMAIL_USER = 'seu.email@seuservidor.com'
EMAIL_PASS = 'sua_senha'
```

---

## Como Usar

1. Certifique-se de que os arquivos de entrada estejam na pasta configurada.
2. Execute o script:
   ```bash
   python script.py
   ```
3. O relatório gerado será salvo no caminho configurado e, caso configurado, será enviado automaticamente por e-mail.

---

## Estrutura do Relatório

### **Índice**
- Aba inicial com hiperlinks para acessar abas específicas do arquivo.
- Categorias organizadas para facilitar a navegação.

### **Abas de Produto**
- Tabelas dinâmicas organizadas por linha de produto, cor e status do veículo.
- Dados sumarizados por categoria para análise detalhada.

---

## Personalização

- **Mapeamento de Categorias**: O dicionário `categorias` pode ser ajustado no script para incluir novas categorias ou modificar as existentes.
- **Estilo Visual**: O script utiliza formatações padrão que podem ser personalizadas diretamente no código para atender às preferências visuais do usuário.

---

## Contribuições

Contribuições são sempre bem-vindas! Caso encontre problemas ou tenha sugestões de melhorias, utilize a aba ["Issues"](https://github.com/seu-usuario/relatorios-automatizados/issues) ou envie um pull request.

---

## Licença

Este projeto está licenciado sob a [Licença MIT](https://opensource.org/licenses/MIT). Consulte o arquivo `LICENSE` para mais informações.

---

## Autor

- **Gabriel Matuck**  
  - **E-mail**: [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

