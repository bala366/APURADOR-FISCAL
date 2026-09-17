# Apurador Fiscal - Lucro Real V5

Projeto desktop Windows em Python/Tkinter.

## Interface V5
- Dashboard vermelho e branco.
- Tela inicial: Apuração Geral Mensal.
- Empresa, CNPJ, regime Lucro Real e competência.
- Saldos anteriores de PIS, COFINS e ICMS.
- Indicadores de XML, entradas, saídas e documentos.
- Resumo de débitos, créditos consumidos e saldo.
- Tabela das NF-e da competência.
- Exportação para Excel.
- Geração e visualização de PDF.
- Apuração por cliente separada, com PDF próprio.
- Itens fiscais com NCM, CFOP e CST.
- Simulador de venda.

## Compilar no GitHub
Suba os arquivos, abra Actions, execute "Compilar Apurador Fiscal Windows V5"
e baixe o Artifact "Apurador-Fiscal-Windows-V5".

## Aviso
Software independente, sem vínculo oficial com a SEFAZ/GO.
A leitura dos XML constitui pré-apuração e o tratamento fiscal deve ser validado.


## Correção V5
A apuração mensal agora mostra explicitamente: Saldo anterior + créditos das compras = crédito disponível; crédito disponível - débitos das vendas = resultado final. Resultado negativo aparece em vermelho como imposto a recolher; positivo em verde como crédito a transportar. O AppId e a pasta de instalação foram mantidos, portanto o instalador V5 atualiza a instalação V4 existente. O banco continua em LOCALAPPDATA e não é apagado pela atualização.
