# TODO — Melhorias futuras

> Lista de ideias e melhorias mapeadas mas que **nao sao prioridade agora**.
> Cada item tem contexto suficiente pra ser pego no futuro sem precisar
> reconstruir o raciocinio.

---

## Em estudo

### [ ] Estudar self-healing real do MSI (auto-reparo de arquivos deletados)

**Contexto:** hoje o MSI **nao** auto-recupera arquivos que o usuario deletar
da pasta de instalacao (`%LocalAppData%\Programs\FeriasAutomacao`). Os
atalhos do StartMenu/Desktop/SendTo sao **non-advertised** (apontam direto
pro `powershell.exe -File gui.ps1`), entao clicar no atalho nao dispara o
`MsiProvideComponent` que verificaria os KeyPaths e reinstalaria o que esta
faltando.

**Workaround atual (opcao D — escolhida):** documentado no README na secao
de Troubleshooting. Se algo sumir, o usuario roda `msiexec /f` ou re-executa
o `.msi` do release pra reparar.

**Pra implementar self-heal de verdade (opcao B do estudo):**

- Criar um **launcher EXE** minusculo (C# console app ou similar) que so
  faz `Start powershell -File gui.ps1`. Compilar single-file, AOT se der.
- No `installer/Product.wxs`, registrar esse EXE como entry point do app.
- Trocar os 3 `<Shortcut>` (StartMenu/Desktop/SendTo) pra `Advertise="yes"`
  e apontar pro EXE (advertised shortcuts so funcionam com EXE, nao com .ps1).
- Validar:
  - Deletar `gui.ps1` -> clicar atalho -> MSI deve reinstalar silenciosamente
  - Deletar o launcher.exe -> clicar atalho -> idem
  - Confirmar que o ProductCode/UpgradeCode continuam compativeis com instalacoes existentes (senao quebra upgrade)

**Quando vale a pena:** se a gente comecar a ver tickets de "o app parou de
abrir do nada". Hoje nao tem evidencia de que isso acontece com frequencia.

**Referencias:**
- WiX docs: https://wixtoolset.org/docs/v4/
- MSI self-healing: `MsiProvideComponent` na MSDN
- Discussao sobre advertised vs non-advertised shortcuts no WiX

---

## Outras ideias (sem prioridade)

<!-- adicionar aqui conforme aparecer -->
