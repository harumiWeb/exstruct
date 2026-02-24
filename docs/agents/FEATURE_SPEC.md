# Feature Spec for AI Agent

## Feature Name

MCP Patch Architecture Refactor (Phase 1)

## èƒŒæ™¯

`src/exstruct/mcp/patch_runner.py` ã¯ 3,500 è¡Œè¶…ã®å˜ä¸€ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ«ã¨ãªã£ã¦ãŠã‚Šã€ä»¥ä¸‹ãŒæ··åœ¨ã—ã¦ã„ã‚‹ã€‚

1. ãƒ‰ãƒ¡ã‚¤ãƒ³ãƒ¢ãƒ‡ãƒ«å®šç¾©ï¼ˆ`PatchOp`, `PatchRequest`, `PatchResult`ï¼‰
2. å…¥åŠ›æ¤œè¨¼ï¼ˆop åˆ¥ãƒãƒªãƒ‡ãƒ¼ã‚·ãƒ§ãƒ³ï¼‰
3. å®Ÿè¡Œåˆ¶å¾¡ï¼ˆengine é¸æŠã€fallbackã€warning é›†ç´„ï¼‰
4. backend å®Ÿè£…ï¼ˆopenpyxl/xlwingsï¼‰
5. å…±é€šãƒ¦ãƒ¼ãƒ†ã‚£ãƒªãƒ†ã‚£ï¼ˆA1 å¤‰æ›ã€path ç«¶åˆå‡¦ç†ã€è‰²å¤‰æ›ï¼‰

ã“ã®çŠ¶æ…‹ã¯ã€ä¿å®ˆæ€§ãƒ»æ‹¡å¼µæ€§ãƒ»ãƒ†ã‚¹ãƒˆå®¹æ˜“æ€§ã‚’ä½ä¸‹ã•ã›ã‚‹ãŸã‚ã€è²¬å‹™åˆ†é›¢ã‚’è¡Œã†ã€‚

## ç›®çš„

1. `patch_runner.py` ã‚’è–„ã„ãƒ•ã‚¡ã‚µãƒ¼ãƒ‰ã¸ç¸®é€€ã™ã‚‹ã€‚
2. patch æ©Ÿèƒ½ã‚’ã€Œãƒ¢ãƒ‡ãƒ«ã€ã€Œæ­£è¦åŒ–/æ¤œè¨¼ã€ã€Œå®Ÿè¡Œåˆ¶å¾¡ã€ã€Œbackend å®Ÿè£…ã€ã«åˆ†é›¢ã™ã‚‹ã€‚
3. `server.py` ã¨ `patch_runner.py` ã«åˆ†æ•£ã—ãŸ patch op æ­£è¦åŒ–ã‚’å…±é€šåŒ–ã™ã‚‹ã€‚
4. é‡è¤‡ãƒ¦ãƒ¼ãƒ†ã‚£ãƒªãƒ†ã‚£ï¼ˆA1ã€å‡ºåŠ› path ç«¶åˆå‡¦ç†ï¼‰ã‚’å…±é€šåŒ–ã™ã‚‹ã€‚
5. å…¬é–‹ API äº’æ›ã‚’ç¶­æŒã—ã¤ã¤ã€æ®µéšçš„ã«ç§»è¡Œå¯èƒ½ãªæ§‹é€ ã«ã™ã‚‹ã€‚

## ã‚¹ã‚³ãƒ¼ãƒ—

### In Scope

1. `src/exstruct/mcp/patch/` é…ä¸‹ã®æ–°è¦ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ«å°å…¥
2. `patch_runner.py` ã®ãƒ•ã‚¡ã‚µãƒ¼ãƒ‰åŒ–
3. patch op æ­£è¦åŒ–ã®å…±é€šåŒ–
4. A1 å¤‰æ›ãƒ»å‡ºåŠ› path è§£æ±ºã®å…±é€šãƒ¦ãƒ¼ãƒ†ã‚£ãƒªãƒ†ã‚£åŒ–
5. æ—¢å­˜ãƒ†ã‚¹ãƒˆã®è²¬å‹™åˆ¥å†é…ç½®ã¨è¿½åŠ 

### Out of Scope

1. æ–°ã—ã„ patch op ã®è¿½åŠ 
2. MCP å¤–éƒ¨ API ã®ä»•æ§˜å¤‰æ›´
3. å¤§è¦æ¨¡ãªãƒ‡ã‚£ãƒ¬ã‚¯ãƒˆãƒªå†ç·¨ï¼ˆ`mcp` å…¨ä½“ã®å†è¨­è¨ˆï¼‰
4. `.xls` ã‚µãƒãƒ¼ãƒˆæ–¹é‡ã®å¤‰æ›´

## ã‚¿ãƒ¼ã‚²ãƒƒãƒˆã‚¢ãƒ¼ã‚­ãƒ†ã‚¯ãƒãƒ£

```text
src/exstruct/mcp/
  patch/
    __init__.py
    types.py
    models.py
    specs.py
    normalize.py
    validate.py
    service.py
    engine/
      base.py
      openpyxl_engine.py
      xlwings_engine.py
    ops/
      common.py
      openpyxl_ops.py
      xlwings_ops.py
  shared/
    a1.py
    output_path.py
```

## ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ«è²¬å‹™

1. `patch/types.py`
   1. `PatchOpType`, `PatchBackend`, `PatchEngine`, `PatchStatus` ç­‰ã®å‹å®šç¾©
2. `patch/models.py`
   1. `PatchOp`, `PatchRequest`, `MakeRequest`, `PatchResult` ã¨ snapshot ç³»ãƒ¢ãƒ‡ãƒ«
3. `patch/specs.py`
   1. op ã”ã¨ã® required/optional/constraints/aliases ã‚’å˜ä¸€ç®¡ç†
4. `patch/normalize.py`
   1. top-level `sheet` é©ç”¨
   2. alias æ­£è¦åŒ–ï¼ˆ`name`, `row`, `col`, `horizontal`, `vertical`, `color` ãªã©ï¼‰
5. `patch/validate.py`
   1. `PatchOp` ã®æ•´åˆæ€§ãƒã‚§ãƒƒã‚¯ï¼ˆspec ãƒ™ãƒ¼ã‚¹ï¼‰
6. `patch/service.py`
   1. `run_patch` / `run_make` ã®ã‚ªãƒ¼ã‚±ã‚¹ãƒˆãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³
   2. engine é¸æŠãƒ»fallbackãƒ»warning/error çµ„ã¿ç«‹ã¦
7. `patch/engine/*`
   1. backend ã”ã¨ã® workbook ç·¨é›†ã¨ä¿å­˜è²¬å‹™
8. `patch/ops/*`
   1. op é©ç”¨ãƒ­ã‚¸ãƒƒã‚¯ï¼ˆbackend åˆ¥ï¼‰
9. `shared/a1.py`
   1. A1 è§£æã€åˆ—å¤‰æ›ã€ç¯„å›²å±•é–‹ã®å…±é€šé–¢æ•°
10. `shared/output_path.py`
   1. `on_conflict`ã€`rename`ã€å‡ºåŠ›å…ˆæ±ºå®šã®å…±é€šé–¢æ•°

## ä¾å­˜ãƒ«ãƒ¼ãƒ«

1. `server.py` ã¯ patch å®Ÿè£…è©³ç´°ã«ä¾å­˜ã—ãªã„ã€‚
2. `op_schema.py` ã¯ `patch_runner.py` ã§ã¯ãªã `patch/specs.py` / `patch/types.py` ã«ä¾å­˜ã™ã‚‹ã€‚
3. `tools.py` ã¯ `patch/service.py` ã¨ `patch/models.py` ã®ã¿ã‚’åˆ©ç”¨ã™ã‚‹ã€‚
4. backend å®Ÿè£…ã¯ `service.py` ã¸ã®é€†ä¾å­˜ã‚’ç¦æ­¢ã™ã‚‹ã€‚
5. å…±é€šé–¢æ•°ã¯ `shared/*` ã«é›†ç´„ã—ã€é‡è¤‡å®Ÿè£…ã‚’ç¦æ­¢ã™ã‚‹ã€‚

## äº’æ›æ€§è¦ä»¶

1. æ—¢å­˜ã®å…¬é–‹ import ã¯ç¶­æŒã™ã‚‹ï¼ˆ`exstruct.mcp.patch_runner` çµŒç”±ã®ä¸»è¦ã‚·ãƒ³ãƒœãƒ«ï¼‰ã€‚
2. MCP tool I/F ã¯å¤‰æ›´ã—ãªã„ï¼ˆå…¥åŠ›ãƒ»å‡ºåŠ› JSON äº’æ›ï¼‰ã€‚
3. æ—¢å­˜ warning/error ãƒ¡ãƒƒã‚»ãƒ¼ã‚¸ã¯å¯èƒ½ãªé™ã‚Šç¶­æŒã™ã‚‹ã€‚
4. æ—¢å­˜ãƒ†ã‚¹ãƒˆï¼ˆ`tests/mcp/test_patch_runner.py` ã»ã‹ï¼‰ã‚’é€šã™ã€‚

## éæ©Ÿèƒ½è¦ä»¶

1. mypy strict: ã‚¨ãƒ©ãƒ¼ 0
2. Ruff (E, W, F, I, B, UP, N, C90): ã‚¨ãƒ©ãƒ¼ 0
3. å¾ªç’°ä¾å­˜ 0
4. 1 ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ« 1 è²¬å‹™ã‚’å„ªå…ˆã—ã€å·¨å¤§é–¢æ•°ã‚’åˆ†å‰²ã™ã‚‹
5. æ–°è¦é–¢æ•°/ã‚¯ãƒ©ã‚¹ã¯ Google ã‚¹ã‚¿ã‚¤ãƒ« docstring ã‚’ä»˜ä¸ã™ã‚‹

## å—ã‘å…¥ã‚Œæ¡ä»¶ï¼ˆAcceptance Criteriaï¼‰

### AC-01 ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ«åˆ†é›¢

1. `patch_runner.py` ãŒãƒ•ã‚¡ã‚µãƒ¼ãƒ‰åŒ–ã•ã‚Œã€å®Ÿè£…è©³ç´°ã®å¤§åŠãŒ `patch/` é…ä¸‹ã¸ç§»å‹•ã—ã¦ã„ã‚‹ã€‚

### AC-02 æ­£è¦åŒ–ä¸€å…ƒåŒ–

1. patch op alias æ­£è¦åŒ–ãƒ­ã‚¸ãƒƒã‚¯ãŒå˜ä¸€ãƒ¢ã‚¸ãƒ¥ãƒ¼ãƒ«ã«é›†ç´„ã•ã‚Œã€`server.py` ã¨ `tools.py` ã‹ã‚‰å†åˆ©ç”¨ã•ã‚Œã‚‹ã€‚

### AC-03 é‡è¤‡å‰Šæ¸›

1. A1 å¤‰æ›ã®é‡è¤‡å®Ÿè£…ãŒé™¤å»ã•ã‚Œã€`shared/a1.py` ã«çµ±ä¸€ã•ã‚Œã‚‹ã€‚
2. å‡ºåŠ› path ç«¶åˆå‡¦ç†ã®é‡è¤‡å®Ÿè£…ãŒé™¤å»ã•ã‚Œã€`shared/output_path.py` ã«çµ±ä¸€ã•ã‚Œã‚‹ã€‚

### AC-04 äº’æ›æ€§ç¶­æŒ

1. æ—¢å­˜ã® MCP ãƒ„ãƒ¼ãƒ«å‘¼ã³å‡ºã—ãŒå›å¸°ãªãå‹•ä½œã™ã‚‹ã€‚
2. æ—¢å­˜ patch/make é–¢é€£ãƒ†ã‚¹ãƒˆãŒé€šéã™ã‚‹ã€‚

### AC-05 å“è³ªã‚²ãƒ¼ãƒˆ

1. `uv run task precommit-run` ãŒæˆåŠŸã™ã‚‹ã€‚

## ãƒ†ã‚¹ãƒˆæ–¹é‡

1. æ—¢å­˜ãƒ†ã‚¹ãƒˆã®å›å¸°ç¢ºèª
   1. `tests/mcp/test_patch_runner.py`
   2. `tests/mcp/test_make_runner.py`
   3. `tests/mcp/test_server.py`
   4. `tests/mcp/test_tools_handlers.py`
2. æ–°è¦ãƒ†ã‚¹ãƒˆè¿½åŠ 
   1. `tests/mcp/patch/test_normalize.py`
   2. `tests/mcp/patch/test_service.py`
   3. `tests/mcp/shared/test_a1.py`
   4. `tests/mcp/shared/test_output_path.py`

## ãƒªã‚¹ã‚¯ã¨å¯¾ç­–

1. ãƒªã‚¹ã‚¯: åˆ†å‰²ä¸­ã« import äº’æ›ãŒå´©ã‚Œã‚‹
   1. å¯¾ç­–: `patch_runner.py` ã§å†ã‚¨ã‚¯ã‚¹ãƒãƒ¼ãƒˆã‚’ç¶­æŒã—ã€æ®µéšç§»è¡Œã™ã‚‹
2. ãƒªã‚¹ã‚¯: warning/error æ–‡è¨€å·®åˆ†ã§ãƒ†ã‚¹ãƒˆãŒå£Šã‚Œã‚‹
   1. å¯¾ç­–: æ—¢å­˜æ–‡è¨€äº’æ›ã‚’ç¶­æŒã—ã€å¿…è¦æ™‚ã¯å·®åˆ†ã‚’æ˜ç¤ºã—ã¦ãƒ†ã‚¹ãƒˆæ›´æ–°ã™ã‚‹
3. ãƒªã‚¹ã‚¯: engine åˆ†é›¢æ™‚ã®æŒ™å‹•å·®
   1. å¯¾ç­–: backend ã”ã¨ã®å›å¸°ãƒ†ã‚¹ãƒˆã‚’å…ˆã«å›ºå®šã—ã¦ã‹ã‚‰ç§»è¡Œã™ã‚‹

---

## Feature Name

MCP Coverage Recovery (Post-Refactor)

## ”wŒi

MCP‘å‹K–ÍƒŠƒtƒ@ƒNƒ^ƒŠƒ“ƒOŒãA‘S‘ÌƒJƒoƒŒƒbƒW‚ª 85% ‚©‚ç 78.24% ‚É’á‰º‚µ‚½B  
`coverage.xml` ‚Ì–¢Àss‚Í 1,654 s‚ÅA‚¤‚¿ `src/exstruct/mcp/*` ‚ª 1,176 si71.1%j‚ğè‚ß‚éB  
åˆö‚ÍˆÈ‰º‚Ì’Ê‚èB

1. `src/exstruct/mcp/patch/internal.py`i59.42%, 806 missj
2. `src/exstruct/mcp/patch/models.py`i77.74%, 181 missj
3. `src/exstruct/mcp/server.py`i70.17%, 71 missj

## –Ú“I

1. ‘S‘ÌƒJƒoƒŒƒbƒW‚ğ 85%ˆÈã‚Ö‰ñ•œ‚µˆÛ‚·‚éB
2. ’á‰º—vˆöƒ‚ƒWƒ…[ƒ‹‚ğƒeƒXƒg‚Å’¼Ú‰ü‘P‚·‚éB
3. `omit` ˆË‘¶‚Å‚ÌŒ©‚©‚¯ã‚Ì‰ñ•œ‚Ís‚í‚È‚¢iÀs•s”\ƒR[ƒh‚ÌÅ¬—áŠO‚Ì‚İ‹–—ejB

## ƒXƒR[ƒv

### In Scope

1. `tests/mcp/patch/*` ‚ÌŠg’£imodels/internal/service’†Sj
2. `tests/mcp/test_server.py` ‚Ì–¢ƒJƒo[•ªŠò’Ç‰Á
3. `tests/mcp/test_sheet_reader.py` / `tests/mcp/test_chunk_reader.py` ‚Ì‹«ŠEƒP[ƒX’Ç‰Á
4. CIƒQ[ƒgİ’è‚Ì‹­‰»i`--cov-fail-under=85` ‚Æ patch coverage 85%j
5. ’Ç‹LƒhƒLƒ…ƒƒ“ƒgi–{d—lEƒ^ƒXƒNE•K—vÅ¬ŒÀ‚ÌƒeƒXƒg—vŒ”½‰fj

### Out of Scope

1. patch op ‚ÌV‹K‹@”\’Ç‰Á
2. ŒöŠJAPId—l‚Ì•ÏX
3. ‘å‹K–ÍƒfƒBƒŒƒNƒgƒŠÄ•Ò

## À‘••ûj

1. `patch/models.py` ‚ÌƒoƒŠƒf[ƒVƒ‡ƒ“•ªŠò‚ğ `pytest.mark.parametrize` ‚Å–Ô—…‚·‚éB
2. `patch/internal.py` ‚Ì openpyxl/xlwings “K—p•ªŠòAƒGƒ‰[•ªŠòA•Û‘¶‰Â”Û•ªŠò‚ğ¬‚³‚Èfixture‚Å–Ô—…‚·‚éB
3. `server.py` ‚Ì alias³‹K‰»EA1ƒp[ƒXEƒGƒ‰[ƒƒbƒZ[ƒWŒo˜H‚ğ–Ô—…‚·‚éB
4. `sheet_reader.py` / `chunk_reader.py` ‚Ì–¢Às‹«ŠEi‹ó“ü—ÍA•s³rangeApagination‹«ŠEj‚ğ’Ç‰Á‚·‚éB
5. CI‚ğu‘S‘Ì85%–¢–‚Å¸”svu•ÏXs85%–¢–‚Å¸”sv‚É‚·‚éB

## ŒöŠJAPI/ƒCƒ“ƒ^[ƒtƒF[ƒX•ÏX

1. PythonŒöŠJAPI‚Ì•ÏX‚Ís‚í‚È‚¢B
2. CIƒCƒ“ƒ^[ƒtƒF[ƒX‚Æ‚µ‚ÄˆÈ‰º‚ğ’Ç‰ÁE•ÏX‚·‚éB
3. ƒeƒXƒgÀsƒRƒ}ƒ“ƒh‚É `--cov-fail-under=85` ‚ğ’Ç‰Á‚·‚éB
4. Codecov `patch` ƒXƒe[ƒ^ƒX–Ú•W‚ğ `85%` ‚Éİ’è‚·‚éB

## ó‚¯“ü‚êŠî€iAcceptance Criteriaj

1. `uv run pytest -m "not com and not render" --cov=exstruct --cov-report=xml --cov-fail-under=85` ‚ª¬Œ÷‚·‚éB
2. `coverage.xml` ‚Ì‘S‘Ì line-rate ‚ª 85%ˆÈã‚Å‚ ‚éB
3. Codecov patch coverage ‚Ì required status ‚ª 85%ˆÈã‚Å‚ ‚éB
4. `patch/internal.py`, `patch/models.py`, `server.py` ‚Ì line-rate ‚ªŒ»ó’l‚æ‚è—LˆÓ‚É‰ü‘P‚µ‚Ä‚¢‚éB
5. `uv run task precommit-run` ‚ª¬Œ÷‚·‚éimypy strict / RuffŠÜ‚ŞjB

## ƒŠƒXƒN‚Æ‘Îô

1. ƒŠƒXƒN: `patch/internal.py` ‚Ì•ªŠò‚ª‘½‚­H”‚ª–c‚ç‚ŞB  
   ‘Îô: ¸”sŒn‚ğ `parametrize` ‰»‚µA1ƒeƒXƒg‚ ‚½‚è‚Ì–Ô—…Œø—¦‚ğÅ‘å‰»‚·‚éB
2. ƒŠƒXƒN: CIƒQ[ƒg‹­‰»‚Åˆê“I‚É¸”s‚ª‘‚¦‚éB  
   ‘Îô: ’iŠK“±“ü‚¹‚¸“¯PR“à‚Å•s‘«ƒeƒXƒg‚ğ“¯“Š“ü‚·‚éB
3. ƒŠƒXƒN: ŒİŠ·ƒŒƒCƒ„[œŠO‚Ö‚ÌŒã‘Ş”»’fB  
   ‘Îô: –{d—l‚Å‚ÍP‹vœŠO‚ğ‹Ö~‚µA•K—v‚ÍŠúŒÀ•t‚«b’è‘[’u‚ğ•Ê³”F‚Æ‚·‚éB
