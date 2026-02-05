# Testes

## Execucao padrao (unit + integracao mock)
```bash
python -m pytest
```

## Stress (carga local)
```bash
set RUN_STRESS=1
python -m pytest -m stress
```

## Live (Outlook/Yube reais)
```bash
set RUN_LIVE_TESTS=1
set RUN_LIVE_RPA=1
set RUN_LIVE_EMAIL=0
set ASO_LIVE_LIMIT=20
set ASO_LIVE_BASE=C:\ASOgui_live
python -m pytest -m live
```
