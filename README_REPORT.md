# Automacao do Reporte

O app agora esta preparado para publicacao sem depender de arquivos `.xlsx` dentro do projeto.

## Como funciona

- o usuario abre o servidor local
- faz upload do arquivo `.xlsx` pela interface
- o dashboard e o `Weekly Report` sao gerados automaticamente
- os arquivos enviados ficam fora da pasta do projeto

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

## Fluxo de uso

1. clique em `Upload And Open Report`
2. escolha um arquivo `.xlsx`
3. o sistema gera o dashboard automaticamente
4. use a aba `Weekly Report` para filtros, exportacao PDF e tabelas executivas

## Arquivos principais

- `server.py`: sobe o servidor, recebe uploads e abre o dashboard
- `generate_report.py`: gera o dashboard HTML
- `Dockerfile`: imagem para deploy no Coolify
- `requirements.txt`: dependencias Python
- `report_output/report.html`: ultimo HTML gerado

## Observacao importante

- o app nao precisa mais que a planilha fique dentro do projeto
- para publicar o repositorio, basta nao incluir arquivos `.xlsx`
