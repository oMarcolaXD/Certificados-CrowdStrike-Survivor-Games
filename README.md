# Gerador de Certificados — CrowdStrike Survivor Games

Aplicativo local para gerar certificados em PDF a partir de templates PowerPoint (.pptx), com interface web rodando direto no seu Mac.

## Requisitos

- **macOS** (qualquer versão recente)
- **Microsoft PowerPoint para Mac** (necessário para converter para PDF)
- **Python 3** (já vem instalado no macOS)

## Como instalar

1. Baixe o arquivo `certificate-app.zip` na seção [Releases](../../releases)
2. Extraia o ZIP em qualquer pasta
3. Abra o Terminal na pasta extraída
4. Execute:

```bash
./iniciar.sh
```

O app abre automaticamente no navegador em `http://localhost:8080`.

## Como usar

### 1. Prepare os templates PowerPoint

Crie um ou dois arquivos `.pptx` com os placeholders abaixo no slide:

| Placeholder | Descrição |
|---|---|
| `{{NOME_COMPLETO}}` | Nome completo do participante |
| `{{PRIMEIRO_NOME}}` | Primeiro nome |
| `{{SOBRENOME}}` | Sobrenome |
| `{{EVENTO}}` | Nome do evento |
| `{{COLOCACAO}}` | Colocação (1st Place / 2nd Place / 3rd Place) |
| `{{DATA}}` | Data do evento |

Você pode ter **dois templates distintos**:
- `modelo_participantes.pptx` — para participantes sem colocação
- `modelo_colocados.pptx` — para 1º, 2º e 3º lugar

### 2. Faça o upload dos templates e da lista

Na interface web:
1. Faça upload do(s) template(s) `.pptx`
2. Adicione os participantes manualmente ou via CSV
3. Clique em **Gerar Certificados**
4. Baixe o `.zip` com todos os PDFs gerados

### Formato do CSV

```
Nome Completo,Evento,Colocacao,Data
João Silva,Survivor Games 2025,,28/05/2025
Ana Lima,Survivor Games 2025,1,28/05/2025
Pedro Costa,Survivor Games 2025,2,28/05/2025
```

A coluna `Colocacao` aceita `1`, `2`, `3` ou vazio (sem colocação).

## Observações

- O app roda **100% offline** — nenhum dado sai do seu computador
- O PowerPoint é aberto em segundo plano apenas para converter o arquivo; não é necessário interagir com ele
- Para encerrar o servidor, pressione `Ctrl+C` no Terminal
