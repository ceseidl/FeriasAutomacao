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

### [ ] Re-adicionar checkbox "Iniciar Ferias Automacao agora" no fim do instalador

**Contexto:** chegou a ser implementado, mas em **4 abordagens diferentes** o
checkbox nunca disparou o `gui.ps1` no clique do Finish:

1. `<CustomAction FileRef="AppGerarRelatorioBat" ExeCommand="">` (tipo 18)
   -> WiX executa o `.bat` direto via CreateProcess, mas `.bat` precisa do
   `cmd.exe` interpretador.
2. `<CustomAction Directory="INSTALLFOLDER" ExeCommand='"...powershell.exe" ... -File "...gui.ps1"'>` com `Return="asyncNoWait"` (tipo 50) -> nao
   disparou.
3. `Order="100"` no `<Publish>` (pra rodar antes do `EndDialog Order=999`
   default do `WixUI_InstallDir`) -> Order ficou correto, mas a CA continuou
   nao sendo invocada.
4. `<CustomAction Directory="INSTALLFOLDER" ExeCommand='cmd.exe /c "...Gerar Relatorio.bat"'>` -> cmd.exe deveria sair logo apos `start /b`, mas o app
   ainda assim nao abriu.

**Tentativa que NAO foi feita ainda:** usar a `WixToolset.Util.wixext` com
`WixSilentExec` ou `WixShellExec` apontando pra um `.lnk` gerado durante
o install. A extensao foi instalada mas nao usei a CA dela direito (o nome
`Wix4WixSilentExec_X64` nao bateu em `<CustomActionRef>`).

**Pra retomar:** o usuario vai pesquisar a forma canonica de "launch app
after install" em WiX 4.0.6 e me passa o approach correto. Sao 3 elementos
que precisam coexistir:

```xml
<Property Id="WIXUI_EXITDIALOGOPTIONALCHECKBOX" Value="1" />
<Property Id="WIXUI_EXITDIALOGOPTIONALCHECKBOXTEXT" Value="Iniciar agora" />
<CustomAction Id="LaunchApplication" ... />
<UI>
  <Publish Dialog="ExitDialog" Control="Finish" Event="DoAction"
           Value="LaunchApplication" Order="100"
           Condition="WIXUI_EXITDIALOGOPTIONALCHECKBOX = 1 and NOT Installed" />
</UI>
```

O ponto que falta acertar e o `<CustomAction>` -- que metodo do MSI roda
um processo externo (powershell+args) DETACHADO de forma que dispare via
DoAction de UI no contexto per-user.

**Referencias:**
- https://wixtoolset.org/docs/v4/
- https://github.com/wixtoolset/issues/issues/7674 (discussion sobre
  WixUI_InstallDir + LaunchApplication em WiX 4)

