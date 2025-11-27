# 📧 Forms Campaign - Sistema de Envio de Convites e Lembretes

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph_API-0078D4.svg)](https://graph.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

Sistema automatizado para envio de convites por email e verificação de respostas do Microsoft Forms, ideal para campanhas de atualização cadastral e coleta de dados corporativos.

## 🎯 Funcionalidades

- ✅ **Envio individualizado** de convites por email via Microsoft Graph API
- ✅ **Verificação automática** de respostas baseada em CSV exportado do Microsoft Forms
- ✅ **Lembretes inteligentes** para destinatários que não responderam
- ✅ **Detecção automática** de domínios corporativos vs. genéricos
- ✅ **Rastreamento completo** em arquivo CSV (tracking.csv)
- ✅ **Privacidade garantida**: cada destinatário recebe email individual
- ✅ **Proteção anti-spam** com intervalos configuráveis entre envios

## 📋 Pré-requisitos

- Python 3.8 ou superior
- Conta Microsoft 365 / Azure AD
- Microsoft Forms (para criação do formulário)

## 🚀 Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/forms-campaign.git
cd forms-campaign
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Configure as credenciais do Azure

Crie um arquivo `.env` na raiz do projeto:

```env
TENANT_ID=seu-tenant-id-aqui
CLIENT_ID=seu-client-id-aqui
```

#### Como obter as credenciais:

1. Acesse o [Portal Azure](https://portal.azure.com)
2. Navegue até **Azure Active Directory** → **App registrations**
3. Clique em **New registration**
4. Configure:
   - **Nome**: Forms Campaign App
   - **Tipo**: Public client/native
5. Copie o **Application (client) ID** → `CLIENT_ID`
6. Copie o **Directory (tenant) ID** → `TENANT_ID`
7. Em **API permissions**, adicione:
   - `User.Read`
   - `Mail.Send`
   - `Files.Read.All`
   - `Sites.Read.All`

## 📁 Estrutura de Arquivos

```
.
├── forms_campaign.py              # Script principal
├── .env                           # Credenciais Azure (não versionado)
├── requirements.txt               # Dependências Python
├── ConvitesFormulario_IMPORT_MIN.csv  # Lista de destinatários (input)
├── respostas_forms.csv            # Respostas do Forms (input)
└── tracking.csv                   # Rastreamento de envios (gerado)
```

### 📝 Formato dos arquivos CSV

**ConvitesFormulario_IMPORT_MIN.csv:**
```csv
Title,Email
Empresa ABC Ltda,contato@empresaabc.com.br
Tech Solutions Inc,suporte@techsolutions.com
```

**respostas_forms.csv:**
> Exportado automaticamente do Microsoft Forms (aba Respostas → Baixar/Exportar)

## 💻 Uso

### Comando 1: Enviar Convites Iniciais

```bash
python forms_campaign.py send \
  --subject "Atualização Cadastral - Reforma Tributária" \
  --form-link "https://forms.cloud.microsoft/r/SEU_FORM_ID"
```

**O que acontece:**
- ✉️ Envia email personalizado para cada destinatário
- 📊 Registra data de envio no `tracking.csv`
- ⏱️ Aguarda 3 segundos entre cada envio (anti-spam)

### Comando 2: Verificar Respostas e Enviar Lembretes

```bash
python forms_campaign.py check \
  --form-link "https://forms.cloud.microsoft/r/SEU_FORM_ID"
```

**O que acontece:**
1. 📥 Lê o arquivo `respostas_forms.csv` exportado do Forms
2. ✅ Marca como "respondido" quem preencheu o formulário
3. 🔔 Envia lembrete **imediato** para quem não respondeu
4. 📝 Atualiza `tracking.csv` com timestamp dos lembretes

**Personalizar assunto do lembrete:**
```bash
python forms_campaign.py check \
  --subject "Lembrete Urgente - Prazo Final" \
  --form-link "https://forms.cloud.microsoft/r/SEU_FORM_ID"
```

## 🧠 Lógica de Validação de Respostas

### Detecção Automática de Domínios Corporativos

O sistema identifica automaticamente domínios corporativos vs. genéricos:

**Domínios Genéricos (não-corporativos):**
- `gmail.com`, `outlook.com`, `hotmail.com`, `live.com`
- `yahoo.com`, `icloud.com`, `bol.com.br`, `uol.com.br`

**Regra de Validação:**

```
SE email exato está no CSV de respostas:
   ✅ Marca como respondido

SENÃO SE domínio é corporativo E alguém desse domínio respondeu:
   ✅ Marca como respondido (validação por domínio)

SENÃO:
   ❌ Não marca como respondido (enviará lembrete)
```

### Exemplo Prático

**Lista de convites:**
- `joao@statomat.com.br`
- `maria@statomat.com.br`
- `pedro@gmail.com`

**CSV de respostas contém:**
- `joao@statomat.com.br`

**Resultado:**
- ✅ João → respondido (email exato)
- ✅ Maria → respondido (domínio corporativo `statomat.com.br` validado)
- ❌ Pedro → **não** respondido (gmail requer email exato)

## ⚙️ Configurações

Edite diretamente no arquivo `forms_campaign.py`:

```python
# Intervalo entre envios (em segundos)
SLEEP_SECONDS_BETWEEN_MAILS = 3  # Recomendado: 2-3 segundos

# Prazo padrão para respostas (dias)
DAYS_DEADLINE = 7

# Domínios genéricos (não-corporativos)
GENERIC_DOMAINS = {
    "gmail.com", "outlook.com", "hotmail.com", 
    "yahoo.com", "icloud.com", ...
}
```

## 📊 Performance

| Quantidade | Intervalo | Tempo Estimado |
|------------|-----------|----------------|
| 100 emails | 3s        | ~5 minutos     |
| 300 emails | 3s        | ~15 minutos    |
| 600 emails | 3s        | ~30 minutos    |

> ⚠️ **Limite Microsoft 365:** 30 emails/minuto (nosso padrão: ~20/min)

## 🔒 Segurança e Privacidade

- 🔐 **Envio individual**: Cada destinatário recebe apenas seu próprio email
- 🚫 **Sem CC/BCC**: Nenhum outro email é visível
- 🏢 **Isolamento de dados**: Fornecedores não veem informações uns dos outros
- 🔑 **Autenticação segura**: MSAL (Microsoft Authentication Library)

## 📧 Template de Email

### Email Inicial

```
Prezados, [Nome da Empresa],

Em virtude da Reforma Tributária em andamento no Brasil, estamos 
atualizando nosso cadastro de fornecedores para garantir a conformidade 
com as novas exigências fiscais.

[Botão: Preencher Formulário]

Link de referência: https://www.gov.br/fazenda/...

Atenciosamente,
Statomat Máquinas Especiais
```

### Email de Lembrete

```
Prezados, [Nome da Empresa],

Este é um lembrete sobre a atualização cadastral solicitada anteriormente.

Até o momento, não identificamos sua resposta...

[Botão: Preencher Formulário Agora]

---
Se você já respondeu ao formulário, por favor desconsidere esta mensagem!
```

## 🛠️ Troubleshooting

### Erro de autenticação

```bash
# Limpe o cache de autenticação
rm -rf ~/.msal_token_cache.bin  # Linux/Mac
del %USERPROFILE%\.msal_token_cache.bin  # Windows
```

### CSV não reconhecido

Certifique-se de que o CSV exportado do Forms contém a coluna:
- `Informe um E-mail para contato` (prioridade)
- OU qualquer coluna com `email` no nome

### Rate limiting (muitos emails)

Aumente o intervalo em `forms_campaign.py`:
```python
SLEEP_SECONDS_BETWEEN_MAILS = 5  # De 3 para 5 segundos
```

## 📝 Tracking CSV

O arquivo `tracking.csv` mantém o histórico completo:

| Campo | Descrição |
|-------|-----------|
| `Title` | Nome da empresa |
| `Email` | Endereço de destino |
| `sent_at_iso` | Data/hora do envio inicial |
| `due_at_iso` | Prazo para resposta (informativo) |
| `responded_at_iso` | Data/hora da resposta |
| `reminder_sent_at_iso` | Data/hora do lembrete |

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fazer fork do projeto
2. Criar uma branch para sua feature (`git checkout -b feature/nova-funcionalidade`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/nova-funcionalidade`)
5. Abrir um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## 👨‍💻 Autor

Desenvolvido para gerenciamento de campanhas corporativas de atualização cadastral.

---

⭐ **Se este projeto foi útil, considere dar uma estrela!**
