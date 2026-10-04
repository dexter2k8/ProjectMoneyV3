---
description: Roda a suite de validação (compilação, lint e testes de execução)
---

Rode a verificação do projeto e trate o resultado antes de responder:

- Se houver qualquer `[FALHA]` ou `error:`, corrija o código e rode a suite de novo até `RESULTADO: ... 0 falhas`.
- Se tudo estiver verde, confirme em uma linha com o resumo `N OK, 0 falhas`.

!`powershell -NoProfile -ExecutionPolicy Bypass -File .\test.ps1`
