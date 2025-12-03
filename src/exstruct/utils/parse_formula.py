from typing import Optional, Dict, List


def _extract_series_args_text(formula: str) -> Optional[str]:
    """
    '=SERIES(' から始まる数式の、最外の括弧内テキストを取り出す。
    文字列/配列/入れ子括弧を考慮して対応する閉じ括弧を探索。
    """
    if not formula:
        return None
    s = formula.strip()
    # 大文字小文字無視で =SERIES を検出
    if not s.upper().startswith("=SERIES"):
        return None
    # 最初の '(' の位置
    try:
        open_idx = s.index("(", s.upper().index("=SERIES"))
    except ValueError:
        return None
    # 対応する閉じ括弧を見つける
    depth_paren = 0
    depth_brace = 0  # 配列 { } は引数分割には影響しないが、内部の区切りを誤検出しないためにトラッキング
    in_str = False
    i = open_idx + 1
    start = i
    while i < len(s):
        ch = s[i]
        if in_str:
            if ch == '"':
                # 連続する "" はエスケープ → 1つの " として扱う
                if i + 1 < len(s) and s[i + 1] == '"':
                    i += 2
                    continue
                else:
                    in_str = False
                    i += 1
                    continue
            else:
                i += 1
                continue
        else:
            if ch == '"':
                in_str = True
                i += 1
                continue
            elif ch == "(":
                depth_paren += 1
            elif ch == ")":
                if depth_paren == 0:
                    # 最外の対応する ')' を発見
                    return s[start:i].strip()
                depth_paren -= 1
            elif ch == "{":
                depth_brace += 1
            elif ch == "}":
                if depth_brace > 0:
                    depth_brace -= 1
            i += 1
    return None  # 対応が見つからない（異常系）


def _split_top_level_args(args_text: str) -> List[str]:
    """
    SERIES の括弧内テキストを、トップレベルの引数に分割。
    区切りは ',' または ';'（.FormulaLocal 互換）。
    文字列/配列/入れ子括弧の内部にある区切りは無視。
    """
    if args_text is None:
        return []
    # 実際に使われている区切りを推定（どちらも混在する可能性はあるが、トップレベルのみ見る）
    # セミコロンが使われていれば、それを優先的に区切りとみなす
    use_semicolon = (";" in args_text) and ("," not in args_text.split('"')[0])
    sep_chars = (";",) if use_semicolon else (",",)
    args: List[str] = []
    buf: List[str] = []
    depth_paren = 0
    depth_brace = 0
    in_str = False
    i = 0
    while i < len(args_text):
        ch = args_text[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(args_text) and args_text[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                else:
                    in_str = False
                    i += 1
                    continue
            else:
                buf.append(ch)
                i += 1
                continue
        else:
            if ch == '"':
                in_str = True
                i += 1
                continue
            elif ch == "(":
                depth_paren += 1
                buf.append(ch)
                i += 1
                continue
            elif ch == ")":
                depth_paren = max(0, depth_paren - 1)
                buf.append(ch)
                i += 1
                continue
            elif ch == "{":
                depth_brace += 1
                buf.append(ch)
                i += 1
                continue
            elif ch == "}":
                depth_brace = max(0, depth_brace - 1)
                buf.append(ch)
                i += 1
                continue
            elif (ch in sep_chars) and depth_paren == 0 and depth_brace == 0:
                args.append("".join(buf).strip())
                buf = []
                i += 1
                continue
            else:
                buf.append(ch)
                i += 1
                continue
    # 最後の引数を追加
    if buf or (args and args_text.endswith(sep_chars)):
        args.append("".join(buf).strip())
    return args


def _unquote_excel_string(s: Optional[str]) -> Optional[str]:
    """
    文字列リテラル（"..."）ならクォートを外し、"" → " に復元。
    それ以外はそのまま返す。
    """
    if s is None:
        return None
    st = s.strip()
    if len(st) >= 2 and st[0] == '"' and st[-1] == '"':
        inner = st[1:-1]
        return inner.replace('""', '"')
    return None  # 参照や関数の場合は文字列リテラルではない


def parse_series_formula(formula: str) -> Optional[Dict[str, Optional[str]]]:
    """
    =SERIES(name, x_range, y_range, plot_order[, bubble_size]) を堅牢に解析。
    - 返り値は従来の name_range/x_range/y_range を必ず含む（取れなければ None）
    - 追加で plot_order / bubble_size_range / name_literal も返す（後方互換性あり）
    """
    args_text = _extract_series_args_text(formula)
    if args_text is None:
        return None
    parts = _split_top_level_args(args_text)
    # 欠損を許容（存在しない場合は None）
    name_part = parts[0].strip() if len(parts) >= 1 and parts[0].strip() != "" else None
    x_part = parts[1].strip() if len(parts) >= 2 and parts[1].strip() != "" else None
    y_part = parts[2].strip() if len(parts) >= 3 and parts[2].strip() != "" else None
    plot_order_part = (
        parts[3].strip() if len(parts) >= 4 and parts[3].strip() != "" else None
    )
    bubble_part = (
        parts[4].strip() if len(parts) >= 5 and parts[4].strip() != "" else None
    )
    # name が "..." 形式ならリテラル、それ以外は参照/式とみなす
    name_literal = _unquote_excel_string(name_part)
    name_range = None if name_literal is not None else name_part
    return {
        "name_range": name_range,
        "x_range": x_part,
        "y_range": y_part,
        # 追加情報（任意）：後方互換性を壊さない
        "plot_order": plot_order_part,
        "bubble_size_range": bubble_part,
        "name_literal": name_literal,
    }
