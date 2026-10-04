# ProjectMoney — instruções do projeto

## Stack

- Lazarus 4.8 / Free Pascal 3.2.2, modo `{$mode objfpc}{$H+}`, alvo win64.
- LCL (formulário em `unitmoney.pas` + `unitmoney.lfm`), dados em `unitDatabase.pas`.
- Banco: SQLite `banks.db`, criado ao lado do `projectMoney.exe`.
  **`sqlite3.dll` (x64, oficial) é obrigatória na raiz** — o FPC carrega em runtime.

## Comandos de verificação

| comando | o que faz |
| --- | --- |
| `powershell -NoProfile -ExecutionPolicy Bypass -File .\lint.ps1` | compilação limpa (`lazbuild -B`) + mensagens do FPC no formato `arquivo(linha,col) severidade: mensagem`; exit 1 se houver erro |
| `powershell -NoProfile -ExecutionPolicy Bypass -File .\test.ps1` | suite completa: lint + testes de execução (criação do banco, idempotência, DLL ausente, crash WER); exit 0 só com tudo verde |

Switches do `test.ps1`: `-SkipLint` (só execução), `-SkipDllTest` (pula o teste de DLL).

No VS Code: `Ctrl+Shift+P` → *Tasks: Run Test Task* roda a suite; `Ctrl+Shift+B` só compila.

## Regra de validação

- **Depois de toda mudança em `*.pas`, `*.lpr`, `*.lfm`, `*.lpi`, `lint.ps1` ou `test.ps1`, rode `.\test.ps1` e corrija qualquer falha antes de reportar a tarefa como concluída.**
- No mínimo, nunca termine uma mudança sem `.\lint.ps1` verde.
- Se a suite falhar por ambiente (ex.: `projectMoney.exe` em uso pelo Windows), rode ao menos o lint e declare claramente o que não foi verificado.
- Não declare sucesso com base apenas em "compilou" — a suite também valida comportamento em execução.

## Convenções do código

- Componentes do form e seus handlers ficam **antes de `private`** (seção publicada): sem isso o streaming do `.lfm` não encontra os membros.
- Nome de handler = `ComponenteEvento` (ex.: `toggleShowControlsClick`) e sempre com o correspondente `On<Event> = ...` no `.lfm`.
- Esquema do banco vive em **um único lugar**: a constante `SqlCreateBanks` em `unitDatabase.pas`. Se mudar a estrutura, atualize `SqlCreateBanks` **e** as checagens de schema do `test.ps1`.
- Sempre que possível, valide mudanças pequenas com o lint e mudanças de comportamento com a suite completa.

## Cuidados

- `banks.db` é **dado do usuário**: o `test.ps1` faz backup e restaura, mas não apague/arritma o arquivo manualmente.
- Não editar artefatos de build (`projectMoney.exe`, `lib\**`, `*.res`, `link*.res`) à mão — gere-os com `lazbuild`.
- `bakns.db` (typo antigo) foi removido; se reaparecer, o arquivo correto é `banks.db`.
- `/.vscode` está no `.gitignore`: `settings.json` e `tasks.json` são locais desta máquina.
  `lint.ps1`, `test.ps1`, `unitDatabase.pas`, `sqlite3.dll` e `AGENTS.md` devem ir para o commit.

## Ambiente de desenvolvimento

- `.vscode/settings.json` configura o OmniPascal (`defaultDevelopmentEnvironment=FreePascal` + `searchPath`
  apontando para `C:\lazarus\lcl\*` e `C:\lazarus\fpc\3.2.2\source\*`). Sem isso, o Error Lens acusa
  erros falsos `Cannot find unit ...`. Após mexer nessas chaves, `Developer: Reload Window`.
