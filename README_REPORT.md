# Automacao do Reporte

O app agora suporta dois fluxos:

- upload de `.xlsx` pela interface
- abertura do dashboard sem workbook, para usar a aba `Flat Time` somente com `.csv`

## Como funciona

- o usuario abre o servidor local
- faz upload do arquivo `.xlsx` pela interface ou abre o dashboard sem workbook
- o dashboard e o `Weekly Report` sao gerados automaticamente
- os arquivos enviados nao ficam salvos no servidor
- a aba `Flat Time` pode funcionar so com `.csv`, mesmo sem workbook carregado

## Rodar o app

Na pasta do projeto:

```bash
python3 server.py
```

Depois abra:

```text
http://127.0.0.1:8000
```

## Deploy no Coolify

Os arquivos de deploy ja foram adicionados:

- `Dockerfile`
- `requirements.txt`
- `.dockerignore`

Configuracao recomendada no Coolify:

1. selecione o repositorio `Weekly-Report`
2. use o `Dockerfile` da raiz
3. mantenha a porta `8000`
4. comando de start:

```text
python server.py
```

O servidor ja aceita `HOST` e `PORT` por variavel de ambiente, entao ele funciona bem em container.

## Endurecimento de seguranca

Para ambientes corporativos mais restritos, o app agora suporta:

- autenticacao basica opcional
- allowlist opcional por IP ou rede
- headers de seguranca HTTP
- processamento do `.xlsx` somente em memoria

Variaveis de ambiente uteis no Coolify:

```text
HOST=0.0.0.0
PORT=8000
BASIC_AUTH_USER=seu_usuario
BASIC_AUTH_PASSWORD=sua_senha_forte
ALLOWED_IPS=187.77.250.164,10.0.0.0/8,192.168.0.0/16
```

Observacoes:

- `BASIC_AUTH_USER` e `BASIC_AUTH_PASSWORD` protegem o site inteiro
- `ALLOWED_IPS` aceita IPs individuais e redes CIDR separadas por virgula
- `/health` continua aberto para o healthcheck do deploy

## Fluxo de uso

1. clique em `Upload And Open Report` e escolha um `.xlsx`
2. ou use `Open Dashboard Without Workbook` para entrar direto no app
3. sem workbook, a aba `Flat Time` continua utilizavel com upload de `.csv`
4. use as abas `Interactive Dashboard`, `Weekly Report` e `Flat Time`

## Arquivos principais

- `server.py`: sobe o servidor, recebe uploads e abre o dashboard
- `generate_report.py`: gera o dashboard HTML
- `Dockerfile`: imagem para deploy no Coolify
- `requirements.txt`: dependencias Python
- `report_output/report.html`: ultimo HTML gerado

## Observacao importante

- o app nao precisa mais que a planilha fique dentro do projeto
- o arquivo enviado e processado somente em memoria
- nenhum `.xlsx` fica embutido no servidor
- em ambientes corporativos, `.csv` normalmente passa com menos restricao porque e texto simples, enquanto `.xlsx` e um pacote Office compactado e costuma ser inspecionado por antivirus, DLP, filtro de upload e politicas anti-malware
