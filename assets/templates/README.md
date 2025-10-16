# Template de Voucher Word

Este diretório contém o template de voucher em formato Word (.docx) que será usado para gerar os vouchers automaticamente.

## Como criar o template:

1. **Crie um documento Word** com o layout desejado para o voucher
2. **Use placeholders** no formato `{{nome_da_variavel}}` onde você quer que os dados sejam inseridos
3. **Salve como .docx** e coloque o arquivo como `voucher_template.docx` neste diretório

## Placeholders disponíveis:

- `{{nome}}` - Nome do hóspede
- `{{tipoAcomodacao}}` - Tipo de acomodação
- `{{observacoes}}` - Observações adicionais
- `{{checkin}}` - Data de check-in (formato dd/mm/aaaa)
- `{{checkout}}` - Data de check-out (formato dd/mm/aaaa)
- `{{dataEmissao}}` - Data de emissão do voucher
- `{{valorDiaria}}` - Valor da diária (formato R$ X,XX)
- `{{valorAntecipado}}` - Valor antecipado (formato R$ X,XX)
- `{{total}}` - Valor total (formato R$ X,XX)
- `{{restante}}` - Valor restante (formato R$ X,XX)
- `{{quantidadeDiarias}}` - Número de diárias
- `{{voucherId}}` - ID único do voucher
- `{{status}}` - Status do voucher (Pago/Pendente)

## Exemplo de template:

```
VOUCHER DE HOSPEDAGEM
ID: {{voucherId}}

Hóspede: {{nome}}
Acomodação: {{tipoAcomodacao}}
Check-in: {{checkin}}
Check-out: {{checkout}}
Diárias: {{quantidadeDiarias}}

Valor da Diária: {{valorDiaria}}
Valor Antecipado: {{valorAntecipado}}
Valor Total: {{total}}
Valor Restante: {{restante}}
Status: {{status}}

Observações: {{observacoes}}

Emitido em: {{dataEmissao}}
```

## Notas importantes:

- O arquivo deve ser salvo como `voucher_template.docx`
- Use formatação normal do Word (negrito, itálico, cores, etc.)
- Os placeholders serão substituídos automaticamente pelos dados do formulário
- Se o template não for encontrado, será usado um template HTML básico

