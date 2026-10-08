import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_styled_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Strict Meiryo UI Font Family across all cells
    FONT_FAMILY = "Meiryo UI"
    
    font_title = Font(name=FONT_FAMILY, size=15, bold=True, color="111111")
    font_subtitle = Font(name=FONT_FAMILY, size=9.5, italic=True, color="444444")
    font_section = Font(name=FONT_FAMILY, size=11, bold=True, color="1A1A1A")
    font_header = Font(name=FONT_FAMILY, size=9.5, bold=True, color="FFFFFF")
    font_header_sub = Font(name=FONT_FAMILY, size=9.5, bold=True, color="FFFFFF")
    font_body = Font(name=FONT_FAMILY, size=9, color="212529")
    font_body_bold = Font(name=FONT_FAMILY, size=9, bold=True, color="212529")
    font_code = Font(name=FONT_FAMILY, size=8.5, color="212529")
    font_badge_danger = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")
    font_badge_success = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")

    # Monotone Color Palette (Strictly Grayscale / Slate / Monochrome)
    fill_header_black = PatternFill(start_color="212529", end_color="212529", fill_type="solid") # Dark Charcoal Black
    fill_header_gray = PatternFill(start_color="495057", end_color="495057", fill_type="solid")  # Slate Gray
    fill_sub_header = PatternFill(start_color="CED4DA", end_color="CED4DA", fill_type="solid")   # Medium Light Gray
    fill_section_bar = PatternFill(start_color="DEE2E6", end_color="DEE2E6", fill_type="solid") # Gray Section Divider
    fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")        # Subtle Light Gray Zebra
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_highlight_alert = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid") # Neutral Gray Accent

    # Borders
    border_thin = Side(style="thin", color="CED4DA")
    border_thick_bottom = Side(style="medium", color="212529")
    border_double_bottom = Side(style="double", color="495057")

    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    header_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
    total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double_bottom)

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # =========================================================================
    # SHEET 1: 01_概要と判定 (Executive Summary & Key Takeaways)
    # =========================================================================
    ws1 = wb.create_sheet(title="01_概要と判定")
    ws1.views.sheetView[0].showGridLines = True

    # Title Banner
    ws1.merge_cells("A1:G1")
    ws1["A1"] = "Amazon Aurora Serverless v2 (MySQL 8.4.7 LTS) パラメータ検証報告書"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:G2")
    ws1["A2"] = "検証対象: innodb_lock_wait_timeout (デフォルト 50秒 → 120秒 への延長設定に関するメリット・デメリット詳細評価)"
    ws1["A2"].font = font_subtitle
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[2].height = 18

    # Executive Verdict Section
    ws1.merge_cells("A4:G4")
    ws1["A4"] = "【結論】 総合アーキテクチャ判定と推奨アクション"
    ws1["A4"].font = font_section
    ws1["A4"].fill = fill_section_bar
    ws1["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[4].height = 24

    ws1.merge_cells("A5:G6")
    ws1["A5"] = (
        "【総合判定：グローバル設定での 120秒 適用は原則『非推奨 (アンチパターン)』】\n"
        "OLTP Web API環境において、行ロック待ち時間を120秒に引き延ばすと、競合発生時の待機コネクションが2.4倍滞留し、"
        "「コネクションプール枯渇」「ALBタイムアウト(60秒)との不整合による幽霊スレッド化」「Aurora Serverless v2の不要なACU過剰スケール(課金高騰)」を引き起こします。\n"
        "【推奨設計】: パラメータグループのグローバル値は 5〜15秒 に短縮し、長時間バッチ処理のみセッション単位 (SET SESSION innodb_lock_wait_timeout = 120;) で個別適用してください。"
    )
    ws1["A5"].font = font_body_bold
    ws1["A5"].fill = fill_highlight_alert
    ws1["A5"].alignment = align_top_left
    ws1["A5"].border = cell_border
    ws1.row_dimensions[5].height = 30
    ws1.row_dimensions[6].height = 30

    # Summary Comparison Matrix
    ws1.cell(row=8, column=1, value="■ 項目別サマリー比較（50秒 vs 120秒）").font = font_section

    headers1 = ["評価軸", "デフォルト (50秒)", "延長変更 (120秒)", "内部動作メカニズムの差異", "Aurora Serverless v2 固有の影響", "総合判定"]
    ws1.row_dimensions[9].height = 26
    for col_idx, h in enumerate(headers1, start=1):
        cell = ws1.cell(row=9, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_black
        cell.alignment = align_center
        cell.border = header_border

    data1 = [
        ("一時的競合の自然解決率", "普通 (50秒待機)", "高い (120秒待機)", "ロック解放待ち許容時間が伸びるため、後続クエリが途中でエラーにならず正常完了する確率は上がる。", "バッチ単体の成功率は上がるが、スレッド滞留とトレードオフ。", "〇 メリット (局所的)"),
        ("コネクションプールの枯渇", "中 (50秒で解放)", "極めて高い (120秒占有)", "詰まったクエリがHikariCPの接続を120秒間抱え込むため、無関係な全SQLがプール待ちで全滅する。", "接続数急増に反応してACUが過剰スケールアップし、高額な課金が発生。", "× 重大なデメリット"),
        ("上位タイムアウトとの整合性", "一部乖離 (ALB 60s)", "完全破綻 (ALB 60s < DB 120s)", "ALBやクライアントが60秒で切断(504)した後も、Aurora内部で誰も待っていないSQLが60秒間待機を継続。", "ユーザーのリロード連打により重複クエリが大量滞留し雪崩現象を誘発。", "× 重大なデメリット"),
        ("デッドロック検知への影響", "影響なし (即時検知)", "影響なし (即時検知)", "InnoDB Deadlock Detectorが即座(ミリ秒単位)にサイクルを検出し片方をロールバックする。", "タイムアウト値を何秒に伸ばしてもデッドロックの挙動・速度は一切不変。", "― 中立 (無関係)"),
        ("ロールバック漏れリスク", "50秒後に発生", "120秒後に発生 (検知遅延)", "innodb_rollback_on_timeout=OFF(標準)のため、タイムアウト文のみ破棄され未完了トランザクションが残る。", "アプリ側で明示的ROLLBACKを行わない場合、長時間ロック保持やデータ不整合を誘発。", "▲ 運用注意"),
        ("アーキテクチャ推奨度", "OLTPには長すぎる", "グローバル適用は危険", "一般的なWeb/OLTPでは 5〜10秒 が適正値。120秒は夜間バッチ・データ移行の局所値。", "グローバル値は縮小し、必要なバッチのみSET SESSIONで120秒に局所化が正解。", "推奨アーキテクチャあり")
    ]

    for r_idx, row_data in enumerate(data1, start=10):
        fill = fill_zebra if r_idx % 2 == 0 else fill_white
        ws1.row_dimensions[r_idx].height = 42
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            if c_idx in [1, 6]:
                cell.alignment = align_center
                if "メリット" in val:
                    cell.font = font_badge_success
                elif "デメリット" in val:
                    cell.font = font_badge_danger
            elif c_idx in [2, 3]:
                cell.alignment = align_center
            else:
                cell.alignment = align_top_left

    # =========================================================================
    # SHEET 2: 02_メリット・デメリット詳細 (Detailed Deep Dive)
    # =========================================================================
    ws2 = wb.create_sheet(title="02_メリット・デメリット詳細")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:F1")
    ws2["A1"] = "innodb_lock_wait_timeout = 120秒 の詳細メリット・デメリット分析"
    ws2["A1"].font = font_title
    ws2.row_dimensions[1].height = 28

    # Table 1: メリット
    ws2.cell(row=3, column=1, value="■ メリット（50秒から120秒へ増やすことで得られる恩恵）").font = font_section
    headers_merit = ["No.", "メリット項目", "詳細メカニズム", "想定されるユースケース", "ビジネス・運用上の価値", "注意点・前提条件"]
    ws2.row_dimensions[4].height = 24
    for c_idx, h in enumerate(headers_merit, start=1):
        cell = ws2.cell(row=4, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_black
        cell.alignment = align_center
        cell.border = header_border

    merits = [
        ("M-1", "長時間バッチ処理の完遂率向上", 
         "大量UPDATE/DELETEや月末締めバッチ等で先行トランザクションと競合した際、待機可能時間が70秒拡大することで、競合相手の終了を待って正常完了できる確率が高まる。", 
         "夜間日次バッチ、月次締め処理、大量データ移行(マイグレーション)", 
         "バッチの異常終了による運用者の夜間緊急呼び出しや、やり直しリカバリ工数の削減。", 
         "本来はバッチ側のコミット単位分割(Chunking)やロック順序統一で解決すべき問題。"),
        ("M-2", "外部API連携遅延の許容", 
         "トランザクション内部で外部決済・物流API等を呼び出しているアーキテクチャにおいて、外部サービスの一時的な通信遅延による即座のロックタイムアウトを回避できる。", 
         "外部決済代行連携、外部SaaS同期処理を含むトランザクション", 
         "外部起因の突発的一時遅延によるトランザクション失敗率の低減。", 
         "トランザクション内で外部通信を行うこと自体が重大な設計アンチパターン。"),
        ("M-3", "リトライ未実装システムの延命", 
         "エラー1205(Lock wait timeout exceeded)発生時に自動再試行(Retry)を行わないレガシーアプリケーションにおいて、エラー画面表示の頻度を抑えることができる。", 
         "リトライ機構が組み込まれていないレガシーアプリ、改修困難なパッケージソフト", 
         "アプリケーションコードに手を入れずに、設定変更だけで一時的にエラー発生率を下げる。", 
         "抜本的な解決ではなく問題の先送りに過ぎず、高負荷時に一気にシステム停止へ発展する。")
    ]

    for r_idx, row in enumerate(merits, start=5):
        ws2.row_dimensions[r_idx].height = 55
        fill = fill_white if r_idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = align_center if c_idx == 1 else align_top_left

    # Table 2: デメリット
    start_demerit_row = 9
    ws2.cell(row=start_demerit_row, column=1, value="■ デメリット（50秒から120秒へ増やすことで発生する深刻な弊害）").font = font_section
    headers_demerit = ["No.", "デメリット項目", "発生メカニズム", "システム・ユーザーへの影響", "Aurora Serverless v2 特有のリスク", "危険度"]
    ws2.row_dimensions[start_demerit_row + 1].height = 24
    for c_idx, h in enumerate(headers_demerit, start=1):
        cell = ws2.cell(row=start_demerit_row + 1, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_gray
        cell.alignment = align_center
        cell.border = header_border

    demerits = [
        ("D-1", "コネクションプールの完全枯渇\n(Cascading Failure)", 
         "待機時間が50sから120sへ2.4倍に伸びるため、待機スレッドがHikariCPやTomcatのコネクションを長時間占有。新規リクエストに割り当てるコネクションがゼロになる。", 
         "ロック競合が起きた1つのテーブルだけでなく、全く無関係なログインや参照系トップページまで全APIが応答不能になりシステム全損。", 
         "DBインスタンスの接続上限だけでなく、RDS Proxyやアプリ側プール枯渇によりクラスターが麻痺する。", "★★★★★\n(壊滅的)"),
        ("D-2", "上位タイムアウト乖離による\n『幽霊待機スレッド』の増殖", 
         "ALB(デフォルト60秒)やブラウザが先にタイムアウトしHTTP 504を返却。ユーザーは再試行でクエリを再送するが、DB内部では旧クエリが残り60秒間無駄にロックを保持し続ける。", 
         "「待っても誰も結果を受け取らない処理」がサーバー資源を浪費し続け、ユーザーの連打によりクエリが倍々ゲームで滞留する。", 
         "無駄な待機スレッドの山積みを『高負荷』と判定し、Auroraが不要にACUを最大までスケールアップしてAWS利用料が高騰する。", "★★★★★\n(高コスト・高負荷)"),
        ("D-3", "障害復旧と原因調査の\n大幅な遅延", 
         "通常なら50秒でエラーになり検知・復旧できる障害が、2分間も沈黙したまま膠着するため、アラート発報やヘルスチェック失敗が大幅に遅れる。", 
         "システムのMTTR(平均復旧時間)が悪化。運用監視ダッシュボードで問題の根本クエリをリアルタイム特定する難易度が急上昇。", 
         "スケールアップ処理と待機スレッドの競合により、オートスケーリングの収束が極端に遅れる。", "★★★★☆\n(運用性悪化)"),
        ("D-4", "MVCC Undoログ肥大化による\n全体性能の劣化", 
         "トランザクションが120秒間オープンされたまま待機状態になると、InnoDBのパージスレッドが不要なUndoログを削除できず滞留する(History List Length増大)。", 
         "他の正常なSELECT文が過去バージョンを探すために古いUndoログを長く辿る必要が生じ、DB全体の参照スループットが低下する。", 
         "Auroraの分散ストレージエンジンにおけるUndoスペース拡張と読み取りレイテンシの上昇を誘発する。", "★★★★☆\n(性能低下)"),
        ("D-5", "innodb_rollback_on_timeout\nデフォルトOFFによるデータ不整合", 
         "120秒経過してエラー1205が発生した際、InnoDBは『失敗したSQL文のみ』をロールバックし、トランザクション全体はロールバックしないのがデフォルト挙動。", 
         "アプリがcatchブロックで明示的ROLLBACKを発行しない場合、中途半端に更新された状態でCOMMITされたり、接続が解放されない。", 
         "120秒待たされた挙句に中途半端なデータ破損が発生するため、原因究明が極めて困難になる。", "★★★★★\n(データ不整合)")
    ]

    for idx, row in enumerate(demerits, start=start_demerit_row + 2):
        ws2.row_dimensions[idx].height = 62
        fill = fill_white if idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws2.cell(row=idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            if c_idx in [1, 6]:
                cell.alignment = align_top_center
                if "★" in val:
                    cell.font = font_badge_danger
            else:
                cell.alignment = align_top_left

    # =========================================================================
    # SHEET 3: 03_小学生向け解説・動作イメージ (Real-World Analogies)
    # =========================================================================
    ws3 = wb.create_sheet(title="03_小学生向け解説・動作イメージ")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:E1")
    ws3["A1"] = "【図解とたとえ話】 小学生でも完全に理解できる！ロックとタイムアウトの仕組み"
    ws3["A1"].font = font_title
    ws3.row_dimensions[1].height = 28

    analogy_sections = [
        ("1. 「公園のすべり台」で考えるデータベース", 
         "登場人物と現実の役割", 
         "・すべり台 ＝ データベースのデータ（1行のレコード）\n"
         "・すべる子供 ＝ トランザクション（SQLを実行するプログラム）\n"
         "・すべり台のルール ＝ 行ロック（ぶつかると危ないので、1人ずつしかすべれない）\n"
         "・階段の行列 ＝ ロック待ち（前の人が終わるのを後ろで待っている状態）\n"
         "・あきらめタイマー ＝ innodb_lock_wait_timeout（いつまで待つかの時間制限）",
         "すべり台のてっぺんで、太郎くんがスマホを見てずっと降りてきません！（未コミット処理）\n"
         "次郎くんはすべりたいので、階段の上でじっと待っています（ロック待ち）。"),
        
        ("2. 「50秒ルール」と「120秒ルール」の違い", 
         "待つ時間の違いが起こす大パニック", 
         "【50秒ルールのとき（デフォルト）】\n"
         "次郎くんは50秒待って太郎くんが降りてこないと、「もういいや！砂場に行こう！」と諦めて階段を降ります。\n"
         "後ろの子供たちも少し待てば階段から出られます。\n\n"
         "【120秒ルールに変えたとき（今回の設定）】\n"
         "次郎くんはなんと「2分間（120秒）」もじっと階段で待ち続けます！\n"
         "その間に、三郎くん、四郎くん、五郎くん…と、後ろにどんどん子供が並んでしまいます。\n"
         "階段はぎゅうぎゅう詰めになり、公園の入り口の外まで大行列が伸びてしまいます！",
         "★大問題：ブランコや砂場で遊びたかった他の子供まで、行列に阻まれて公園に入れなくなります（これが『コネクションプールの枯渇』です！）。"),

        ("3. 「親切な警備員さん（Aurora Serverless v2）」の大失敗", 
         "待つだけでクラウド料金が高くなるワケ", 
         "公園には優秀な警備員さん（Aurora Serverless v2）がいます。\n"
         "警備員さんは「大変だ！ 公園の入り口に子供が何百人も溢れているぞ！ 公園が狭すぎるんだ！」と勘違いします。\n"
         "そして、慌てて隣の土地を急遽高いお金で借りて、公園を2倍、4倍に広げてしまいます（ACUの自動スケールアップ）。\n\n"
         "本当は、太郎くん1人がすべり台でスマホを見ているだけなのに、\n"
         "みんなが120秒も待つせいで大行列ができ、無駄に大金（AWSの費用）を払うことになってしまうのです！",
         "★教訓：待つ時間を長くすると、問題が解決するどころか、無駄なお金まで取られてしまいます。"),

        ("4. 「デッドロック」と「ロック待ち」の違い", 
         "よくある勘違いをスッキリ整理", 
         "【普通のロック待ち（片想い待ち）】\n"
         "・太郎くんが降りるのを次郎くんが待っているだけ。\n"
         "・太郎くんがすべり降りれば、次郎くんもすべれます。\n"
         "・これには『あきらめタイマー（50秒や120秒）』が使われます。\n\n"
         "【デッドロック（お見合い・にらめっこ）】\n"
         "・太郎くん「次郎くんがスコップを貸してくれたら、すべり台を降りるよ」\n"
         "・次郎くん「太郎くんがすべり台を降りてくれたら、スコップを貸すよ」\n"
         "・これは100年待っても絶対に解決しません！\n"
         "・だから監視カメラ（デッドロック検出器）がついていて、この状態になった瞬間に『0.001秒』で片方に強制終了を命じます。\n"
         "・※つまり、デッドロックの解決に50秒も120秒も全く関係ありません！",
         "★まとめ：120秒に伸ばしても、デッドロックは1ミリ秒も変わりません！")
    ]

    r_cur = 3
    for title, sub, content, note in analogy_sections:
        ws3.merge_cells(start_row=r_cur, start_column=1, end_row=r_cur, end_column=5)
        ws3.cell(row=r_cur, column=1, value=title).font = font_section
        ws3.cell(row=r_cur, column=1).fill = fill_section_bar
        ws3.row_dimensions[r_cur].height = 24
        r_cur += 1

        ws3.cell(row=r_cur, column=1, value=sub).font = font_header_sub
        ws3.cell(row=r_cur, column=1).alignment = align_center
        ws3.cell(row=r_cur, column=1).fill = fill_header_gray
        ws3.cell(row=r_cur, column=1).border = cell_border

        ws3.merge_cells(start_row=r_cur, start_column=2, end_row=r_cur, end_column=4)
        c_content = ws3.cell(row=r_cur, column=2, value=content)
        c_content.font = font_body
        c_content.alignment = align_top_left
        c_content.border = cell_border

        c_note = ws3.cell(row=r_cur, column=5, value=note)
        c_note.font = font_body_bold
        c_note.alignment = align_top_left
        c_note.fill = fill_highlight_alert
        c_note.border = cell_border

        ws3.row_dimensions[r_cur].height = 110
        r_cur += 2

    # =========================================================================
    # SHEET 4: 04_歴史背景と技術動作原理 (Engine Internals & Architecture)
    # =========================================================================
    ws4 = wb.create_sheet(title="04_歴史背景と技術動作原理")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:F1")
    ws4["A1"] = "技術的動作原理と Amazon Aurora Serverless v2 (MySQL 8.4) 内部動作"
    ws4["A1"].font = font_title
    ws4.row_dimensions[1].height = 28

    ws4.cell(row=3, column=1, value="■ 1. 誕生の歴史と背景（なぜタイムアウトが必要になったのか）").font = font_section
    history_headers = ["時代・バージョン", "採用アーキテクチャ", "ロックの粒度と課題", "タイムアウト機構の役割と進化"]
    ws4.row_dimensions[4].height = 24
    for c_idx, h in enumerate(history_headers, start=1):
        cell = ws4.cell(row=4, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_black
        cell.alignment = align_center
        cell.border = header_border

    hist_data = [
        ("1990年代 (MySQL 3.x)", "MyISAM ストレージエンジン", "テーブルロック (Table Lock)\n1行の書き込みで表全体を施錠。読込は速いが書込の並行性が著しく低く、Webの大規模化で破綻。", "タイムアウト以前にテーブル全体の空き待ちキューに並ぶ構造。競合時の全体停止が日常茶飯事。"),
        ("2000年代 (MySQL 4.x〜5.x)", "InnoDB の採用と MVCC", "行レベルロック (Row-Level Lock) + MVCC\n更新行のみをロックし、読込はUndoログから過去版を参照。同時実行性が劇的に向上。", "「未コミット放置」「通信途絶」「バグによるハング」で永久にロックが解放されない事態を防ぐため、安全弁として50秒タイムアウトが誕生。"),
        ("2010年代 (Aurora 登場)", "Aurora 分散ストレージ", "コンピュート・ストレージ分離\nInnoDBのロック機構はコンピュートメモリ内に保持し、ストレージは6重化非同期レプリケーション。", "コンピュート障害時のフェイルオーバー時に、ロック情報も新インスタンスへ素早く引き継ぐ必要性が生じる。"),
        ("2024年〜 (MySQL 8.4 LTS)", "Performance Schema 完全統合", "data_locks / data_lock_waits への統一\n従来の innodb_locks は完全廃止。監視負荷を極小化しつつロックグラフを追跡可能に。", "マイクロサービス化・サーバレス化が進み、50秒すら長すぎるため、5〜10秒への短縮やNOWAIT/SKIP LOCKEDが主流へ。")
    ]

    for r_idx, row in enumerate(hist_data, start=5):
        ws4.row_dimensions[r_idx].height = 50
        fill = fill_white if r_idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = align_center if c_idx == 1 else align_top_left

    # Aurora Serverless v2 特有挙動
    start_v2_row = 11
    ws4.cell(row=start_v2_row, column=1, value="■ 2. Aurora Serverless v2 特有の内部動作と悪循環メカニズム").font = font_section
    v2_headers = ["フェーズ", "影響レイヤー", "発生する物理現象", "内部メトリクス・挙動", "システムへの最終影響"]
    ws4.row_dimensions[start_v2_row + 1].height = 24
    for c_idx, h in enumerate(v2_headers, start=1):
        cell = ws4.cell(row=start_v2_row + 1, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_gray
        cell.alignment = align_center
        cell.border = header_border

    v2_data = [
        ("Phase 1", "InnoDB Lock Mgr", "先行トランザクションが更新行のXロックを保持したまま停滞", "data_lock_waits にレコード生成、スレッドが sleep 状態へ", "後続トランザクションが待機タイマー(120秒)を開始"),
        ("Phase 2", "App Connection Pool", "120秒待機のため HikariCP コネクションが返却されず保持継続", "ActiveConnections / BlockedTransactions が急増", "新規リクエストが「コネクション取得待ち」で詰まり始める"),
        ("Phase 3", "ALB / Client", "60秒経過時点で ALB が HTTP 504 Gateway Timeout を返却", "HTTP 5xx エラー急増、ブラウザ側切断", "ユーザーがリロードを連打し、新たな待機クエリがDBへ殺到"),
        ("Phase 4", "Aurora Serverless v2", "Threads_connected急増とスレッドスタックメモリ逼迫を検知", "ServerlessDatabaseCapacity (ACU) が自動スケールアップ", "何も仕事をしていない待機スレッドのために無駄にACUが倍増・課金増"),
        ("Phase 5", "InnoDB Undo Purge", "最古の待機トランザクションにより Undo パージが停止", "History List Length (HLL) が急激に伸長", "参照系クエリ(SELECT)まで過去バージョン走査で重くなり全面ダウン")
    ]

    for idx, row in enumerate(v2_data, start=start_v2_row + 2):
        ws4.row_dimensions[idx].height = 42
        fill = fill_white if idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws4.cell(row=idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = align_center if c_idx in [1, 2] else align_top_left

    # =========================================================================
    # SHEET 5: 05_推奨設計と運用SQL (Best Practices & Architecture)
    # =========================================================================
    ws5 = wb.create_sheet(title="05_推奨設計と運用SQL")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "推奨アーキテクチャ設計・設定値および最新MySQL 8.4調査SQL"
    ws5["A1"].font = font_title
    ws5.row_dimensions[1].height = 28

    ws5.cell(row=3, column=1, value="■ 1. タイムアウト階層の黄金比（レイヤードタイムアウト設計）").font = font_section
    to_headers = ["システム階層", "推奨タイムアウト値", "現在の設定(120s適用時)", "設計上の重要原則", "判定"]
    ws5.row_dimensions[4].height = 24
    for c_idx, h in enumerate(to_headers, start=1):
        cell = ws5.cell(row=4, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_black
        cell.alignment = align_center
        cell.border = header_border

    to_data = [
        ("CloudFront / CDN", "30 秒 〜 60 秒", "30 秒 〜 60 秒", "最上位層。ユーザー体験を保護するため最も短く設定", "OK"),
        ("ALB / API Gateway", "30 秒 〜 60 秒", "60 秒", "クライアント切断の境界線。ここよりDBが長いと幽霊スレッド化する", "乖離発生中"),
        ("Application (ECS/EKS)", "15 秒 〜 30 秒", "30 秒 〜 60 秒", "Webフレームワークのリクエストタイムアウト", "乖離発生中"),
        ("Connection Pool (HikariCP)", "5 秒 〜 10 秒 (connectionTimeout)", "30 秒", "プールからの接続取得待ち時間。長すぎるとスレッド枯渇を誘発", "注意"),
        ("DB クエリタイムアウト", "5 秒 〜 10 秒 (Statement Timeout)", "無制限 または 60秒", "単一SQLの最大実行時間", "注意"),
        ("DB innodb_lock_wait_timeout (Global)", "5 秒 〜 15 秒", "120 秒 (★問題箇所)", "『最下層のロック待ちは上位層より圧倒的に短くする』のが鉄則！", "× 逆転現象 (危険)"),
        ("DB innodb_lock_wait_timeout (Session)", "120 秒 〜 300 秒", "未適用", "夜間バッチ・データ移行セッションのみ明示的に延長する", "◎ 推奨方式")
    ]

    for r_idx, row in enumerate(to_data, start=5):
        ws5.row_dimensions[r_idx].height = 36
        fill = fill_white if r_idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws5.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = cell_border
            if c_idx in [1, 2, 3, 5]:
                cell.alignment = align_center
                if "×" in val:
                    cell.font = font_badge_danger
                elif "◎" in val:
                    cell.font = font_badge_success
            else:
                cell.alignment = align_top_left

    # 実践SQL集
    start_sql_row = 14
    ws5.cell(row=start_sql_row, column=1, value="■ 2. 実践運用クエリ（MySQL 8.4 LTS / Aurora Serverless v2 対応）").font = font_section
    sql_headers = ["用途", "対象機能", "実行SQLコード (MySQL 8.4 LTS 準拠)", "解説と運用ポイント"]
    ws5.row_dimensions[start_sql_row + 1].height = 24
    for c_idx, h in enumerate(sql_headers, start=1):
        cell = ws5.cell(row=start_sql_row + 1, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_gray
        cell.alignment = align_center
        cell.border = header_border

    sql_data = [
        ("推奨設定", "セッション単位のタイムアウト変更", 
         "-- バッチ処理の接続直後に実行\nSET SESSION innodb_lock_wait_timeout = 120;\n\n-- 目的の大量更新処理を実行\nUPDATE orders SET batch_status = 1 WHERE ...;\n\n-- 処理完了後にコミット\nCOMMIT;", 
         "グローバル設定は安全な短時間(5〜15秒)に保ち、長時間待機が必要なバッチ接続だけを安全に個別延長する。"),
        
        ("最新構文", "待機ゼロ構文 (NOWAIT)", 
         "SELECT * FROM orders \nWHERE status = 'UNPROCESSED' \nLIMIT 1 FOR UPDATE NOWAIT;", 
         "他がロック中なら1ミリ秒も待たずに Error 3572 を返す。待機スレッドを一切作らずに次のタスクへ移れる。"),

        ("最新構文", "スキップ構文 (SKIP LOCKED)", 
         "SELECT * FROM task_queue \nWHERE status = 'QUEUED' \nLIMIT 10 FOR UPDATE SKIP LOCKED;", 
         "ロック中の行を自動的に除外して空いている行だけを即時取得。ワーカー並列処理でロック待ちを完全撲滅できる。"),

        ("リアルタイム監視", "ロック待ちブロッキング関係の特定\n(MySQL 8.4 対応)", 
         "SELECT \n  r.trx_id AS waiting_trx, r.trx_mysql_thread_id AS waiting_tid,\n  TIMESTAMPDIFF(SECOND, r.trx_wait_started, NOW()) AS wait_sec,\n  b.trx_id AS blocking_trx, b.trx_mysql_thread_id AS blocking_tid,\n  b.trx_query AS blocking_query\nFROM performance_schema.data_lock_waits w\nJOIN information_schema.innodb_trx r ON r.trx_id = w.requesting_engine_transaction_id\nJOIN information_schema.innodb_trx b ON b.trx_id = w.blocking_engine_transaction_id;", 
         "MySQL 8.4で完全廃止された innodb_locks に代わり、performance_schema.data_lock_waits を結合して犯人クエリを特定。")
    ]

    for idx, row in enumerate(sql_data, start=start_sql_row + 2):
        ws5.row_dimensions[idx].height = 68
        fill = fill_white if idx % 2 != 0 else fill_zebra
        for c_idx, val in enumerate(row, start=1):
            cell = ws5.cell(row=idx, column=c_idx, value=val)
            cell.fill = fill
            cell.border = cell_border
            if c_idx == 1:
                cell.font = font_body_bold
                cell.alignment = align_center
            elif c_idx == 3:
                cell.font = font_code
                cell.alignment = align_top_left
            else:
                cell.font = font_body
                cell.alignment = align_top_left

    # Column Widths
    col_widths = {
        "01_概要と判定": {1: 24, 2: 24, 3: 24, 4: 44, 5: 38, 6: 22, 7: 15},
        "02_メリット・デメリット詳細": {1: 8, 2: 28, 3: 44, 4: 36, 5: 36, 6: 18},
        "03_小学生向け解説・動作イメージ": {1: 22, 2: 26, 3: 50, 4: 20, 5: 34},
        "04_歴史背景と技術動作原理": {1: 20, 2: 24, 3: 42, 4: 42, 5: 34, 6: 20},
        "05_推奨設計と運用SQL": {1: 26, 2: 24, 3: 44, 4: 40, 5: 18, 6: 15}
    }

    for sheet_name, widths in col_widths.items():
        ws = wb[sheet_name]
        for col_idx, width in widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

    # Save
    output_path = r"C:\Users\taka_\Agy\Aurora_InnoDB_Lock_Wait_Agy\Aurora_MySQL_innodb_lock_wait_timeout_Report.xlsx"
    wb.save(output_path)
    print(f"Successfully generated: {output_path}")

if __name__ == "__main__":
    create_styled_workbook()
