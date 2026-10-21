# Planilha-de-Investimentos-DIO.me
Simulador de Investimentos em Fundos Imobiliários (FIIs)

Uma planilha em Excel feita para simular a formação de patrimônio por meio de aportes mensais e estimar a renda mensal potencial, além de distribuir o aporte entre seis categorias de Fundos de Investimento Imobiliário (FIIs), de acordo com o perfil selecionado.

**Arquivo:** `Simulador_Investimentos_FIIs.xlsx`  
**Abas:** `Simulador` (interface e projeções) e `Apoio` (percentuais por perfil).

> **Atenção:** é uma ferramenta educacional de simulação. As taxas e os percentuais são premissas ilustrativas, não garantia de rentabilidade, promessa de dividendos nem recomendação personalizada de investimentos.

## 1. Quais perguntas a ferramenta responde?

| Pergunta | Onde encontrar | Como funciona |
|---|---|---|
| **Quanto investir por mês?** | `Simulador!D15` | Campo editável: aporte mensal informado pelo usuário (inicialmente R$ 500,00). |
| **Por quantos anos?** | `Simulador!D16` | Campo editável: prazo da simulação (inicialmente 5 anos). |
| **Qual é a taxa de rendimento mensal?** | `Simulador!D17` | Taxa aplicada à projeção; referencia a configuração em `D8` (inicialmente **0,6% ao mês**). |
| **Quanto patrimônio será acumulado?** | `Simulador!D18` | Resultado calculado automaticamente pela função financeira `VF` (armazenada como `FV` no arquivo `.xlsx`). |
| **Quanto receberá de dividendos por mês?** | `Simulador!D19` | Estimativa calculada como patrimônio acumulado × taxa mensal. |

Outros campos úteis: `D20` contém o **total aportado** e `D21` mostra o **rendimento acumulado estimado** (patrimônio menos aportes). O perfil de investidor é selecionado em `D22`.

### Configurações e comparações

- **Salário-base:** `Simulador!D7` = **R$ 5.000,00**.
- **Taxa inicial de rendimento da carteira:** `Simulador!D8` = **0,6% ao mês**.
- **Percentual sugerido para aporte:** `Simulador!D10` = **30% do salário**.
- **Aporte sugerido:** `Simulador!D9` = `=D7*D10`, ou **R$ 1.500,00/mês** na configuração inicial.
- **Simulação alternativa com aporte sugerido:** `Simulador!H15:I21`, usando o mesmo prazo e taxa da simulação principal.
- **Cenários de 2, 5, 10, 20 e 30 anos:** linhas `28:32`; projeção com aporte informado em `B:F` e com aporte sugerido em `H:K`.
- **Distribuição do aporte por categoria de FII:** `Simulador!B39:E46`, atualizada conforme o perfil escolhido em `D22`.

## 2. Como entram as funções `VF` e `PROCV`?

### `VF`: patrimônio futuro

A função **VF** (valor futuro; **FV** no padrão interno do arquivo Excel) calcula o patrimônio final de uma sequência de aportes mensais com juros compostos.

**Fórmula da simulação principal (`D18`):**

```excel
=VF(taxa_mensal;D16*12;-aporte;0;0)
```

No arquivo `.xlsx`, a mesma fórmula pode ser exibida ou armazenada com o nome internacional e separadores correspondentes:

```excel
=FV(taxa_mensal,D16*12,-aporte,0,0)
```

- `taxa_mensal`: rentabilidade por mês;
- `D16*12`: converte o prazo de anos para meses;
- `-aporte`: pagamento mensal informado como saída de caixa, para que o valor futuro apareça positivo;
- `0`: ausência de saldo inicial;
- `0`: aporte realizado ao **fim de cada mês**.

A mesma lógica aparece em `E28:E32` para os prazos fixos de 2, 5, 10, 20 e 30 anos, e em `I18` e `J28:J32` na simulação com aporte sugerido.

**Dividendos estimados (`D19`):**

```excel
=D18*taxa_mensal
```

A fórmula supõe que a taxa mensal também representa a taxa de distribuição de renda sobre o patrimônio final. **Na prática, valorização total e rendimento em dividendos são conceitos diferentes**, e os pagamentos dos FIIs podem variar. Portanto, este resultado deve ser interpretado como aproximação didática.

### `PROCV`: percentuais de alocação por perfil

A aba `Apoio` contém uma tabela nas células `B3:E20` com 18 combinações: **3 perfis × 6 tipos de FII**.

| Coluna | Informação | Exemplo |
|---|---|---|
| `B` | Chave composta (perfil + hífen + tipo) | `Moderado-PAPEL` |
| `C` | Perfil de investimento | `Moderado` |
| `D` | Tipo de fundo | `PAPEL` |
| `E` | Percentual da carteira | `32%` |

A fórmula para recuperar o percentual de **PAPEL** (`Simulador!C40`) é:

```excel
=PROCV($D$22&"-"&B40;Apoio!$B$3:$E$20;4;FALSO)
```

A fórmula equivalente armazenada no `.xlsx` utiliza `VLOOKUP(...,FALSE)`.

O Excel concatena o perfil selecionado em `$D$22` com a categoria em `B40` para formar, por exemplo, `Moderado-PAPEL`. Em seguida, o `PROCV` busca uma correspondência **exata** na primeira coluna do intervalo de apoio e devolve o percentual da **4ª coluna** (`E`).

O mesmo cálculo é repetido em `C40:C45`. Já `D40:D45` distribui o dinheiro usando `=aporte*C40` (e assim sucessivamente). A linha 46 soma as porcentagens e os valores, permitindo conferir se o total é **100%** e se equivale ao aporte informado.

## 3. Intervalos nomeados

Para facilitar a leitura das fórmulas e evitar referências frágeis, a planilha usa os seguintes nomes definidos no escopo da pasta de trabalho:

| Nome | Célula referenciada | Significado |
|---|---|---|
| `aporte` | `Simulador!$D$15` | Aporte mensal informado pelo usuário. |
| `taxa_mensal` | `Simulador!$D$17` | Rendimento mensal usado nos cálculos. |
| `salario` | `Simulador!$D$7` | Salário mensal de referência, inicialmente R$ 5.000,00. |
| `aporte_sugerido` | `Simulador!$D$9` | Aporte calculado pelo percentual do salário (30%, inicialmente). |

**Exemplo de vantagem:** `=VF(taxa_mensal;D16*12;-aporte;0;0)` informa claramente que a taxa e o aporte vêm de campos específicos, em vez de exigir que o leitor memorize endereços de células. Os nomes podem ser consultados no Excel em **Fórmulas → Gerenciador de Nomes**.

> Nem todos os nomes definidos precisam aparecer em todas as fórmulas. Por exemplo, `salario` foi definido, mas a fórmula atual de `D9` usa diretamente `D7*D10`.

## 4. Percentuais por perfil e origem dos dados

A distribuição considera seis categorias de FIIs. Os percentuais foram **reproduzidos da tabela de apoio do arquivo de exemplo fornecido**, na aba original `Planilha2`. Na nova planilha, estão registrados em `Apoio!B3:E20`.

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---:|---:|---:|
| Papel | 30% | 32% | 50% |
| Tijolo | 50% | 35% | 10% |
| Híbridos | 10% | 8% | 5% |
| Fundos de fundos (FOFs) | 10% | 5% | 5% |
| Desenvolvimento | 0% | 10% | 20% |
| Hotelarias | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

**Origem e justificativa:** são os mesmos pesos presentes no arquivo-modelo anexado pelo usuário; **não foram calculados a partir de rentabilidades históricas, estudos de risco ou orientação de especialistas**. Foram mantidos para demonstrar validação de dados, chave composta e `PROCV`, e cada perfil soma 100%. A classificação dos pesos como “Conservador”, “Moderado” ou “Agressivo” também pertence ao modelo de referência e não deve ser tomada como avaliação técnica independente do risco dos FIIs.

Para alterar a divisão, edite os percentuais da coluna `E` na aba `Apoio`, mantendo **100% por perfil**. Na aba `Simulador`, a lista suspensa em `D22` oferece as opções `Conservador`, `Moderado` e `Agressivo`.

## 5. Como usar

1. Abra `Simulador_Investimentos_FIIs.xlsx` no Microsoft Excel.
2. Na aba `Simulador`, atualize o salário (`D7`), a rentabilidade mensal estimada (`D8`) e, se desejar, o percentual sugerido (`D10`).
3. Informe o aporte mensal (`D15`) e o prazo desejado (`D16`).
4. Consulte o patrimônio previsto (`D18`), os dividendos estimados (`D19`) e os demais resultados.
5. Selecione o perfil de investidor em `D22` para atualizar automaticamente os percentuais e os valores da distribuição.
6. Compare os cenários de longo prazo e a simulação alternativa com aporte de 30% do salário.

## 6. Premissas e limitações

- Os aportes são feitos ao **final de cada mês**.
- A taxa mensal é considerada **constante** durante toda a projeção, com capitalização composta e reinvestimento implícito dos rendimentos.
- Não são modelados inflação, impostos, corretagem, administração, vacância, inadimplência, oscilação das cotas ou alterações da política de distribuição dos fundos.
- Os resultados são apresentados em **valores nominais** e não constituem uma previsão de mercado.
- O rendimento total de uma carteira de FIIs não equivale necessariamente ao valor distribuído em dividendos.
- Alterações manuais na tabela de apoio devem manter as chaves e os nomes de categorias consistentes, para que o `PROCV` encontre a correspondência exata.

---

**Resumo técnico:** o `VF` projeta o patrimônio; os intervalos nomeados tornam as fórmulas legíveis; a lista suspensa seleciona o perfil; o `PROCV` busca os percentuais na aba `Apoio`; e a multiplicação desses percentuais pelo aporte gera a divisão mensal da carteira.
