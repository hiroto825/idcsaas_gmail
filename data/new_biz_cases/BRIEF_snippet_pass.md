# 企業発 新規事業事例 候補洗い出し（スニペット段階）ブリーフ 2026-10-30

## 前提
- 収録基準・判定ルールは `docs/new_business_case_collection_prompt.md` と同じ（基本条件3つ＋A：成功実績 または B：独自性）。
- この環境では企業サイト等の本文を開けない（WebFetchは全ホスト拒否）。WebSearch の結果（タイトル・URL・抜粋）だけが使える。
- そのため **判定は全件「保留」**、保留理由に「本文未確認」を含める。A・Bは「暫定」評価として書く。
- 対象期間：2026-01-01〜2026-10-30。候補には、2026年の日付がスニペットで確認できる出典を少なくとも1つ付ける。日付を確認できない場合は `period_confirmed: false`。

## やること（担当企業ごと）
1. WebSearch で、その企業グループ（子会社含む）の2026年の新規事業・新サービスの有料提供開始、事業の売上・契約数の伸び、黒字化などを探す。例：「<企業名> 新規事業 2026」「<企業名> 販売開始 2026」「<企業名> 決算説明 新規事業 2026」「<企業名> 子会社 サービス 提供開始 2026」。1社あたり3〜5回程度の検索。
2. 基本条件（既存企業発／具体的な商品・サービス／有料提供開始）に当てはまりそうな候補を挙げる。実証段階・組織設立だけ・販売予定のみ・再販売・買収先の紹介・単なる機能追加は除外候補として記録。
3. 候補は1社あたり0〜6件程度。水増ししない。見つからなければ0件で、確認した範囲を書く。
4. 証券コード・業種・公式サイトは検索で確認できたものだけ書く。確認できなければ `"未確認"`。

## 出力：`data/new_biz_cases/raw/<batch>.json`
```json
{
  "batch": "batch1",
  "started_at": "ISO8601", "finished_at": "ISO8601", "search_count": 0,
  "companies": [
    {"company": "", "code": "", "industry": "東証33業種", "website": "",
     "status": "一部調査", "checked_scope": "検索した観点・クエリの要約",
     "candidates": 0, "excluded": 0, "remaining": "本文で確認すべきこと"}
  ],
  "candidates": [
    {
      "company": "事業主体", "parent": "親会社", "code": "親会社の証券コード",
      "entity_type": "自社 | 子会社 | スピンアウト | 共同事業",
      "business_name": "", "country": "JP",
      "launch_date": "提供開始時期（不明なら未確認）",
      "customer_problem": "対象顧客・課題",
      "offering_value": "提供内容・価値",
      "revenue_model": "収益モデル（未確認可）",
      "origin_assets_domain": "従来の事業・活用資産・新たな展開領域",
      "paid_evidence": "有料提供の根拠（スニペットの記載）／未確認",
      "a_success": {"provisional": "該当見込み | 参考実績あり | 未確認", "aspects": [], "facts": "数値・期間・比較時点（スニペット記載のみ）"},
      "b_uniqueness": {"provisional": "該当見込み | 未確認 | 非該当見込み", "aspects": [], "comparison": "", "difference": "", "customer_effect": ""},
      "judgement": "保留 | 除外",
      "reason": "保留・除外理由（必ず『本文未確認』を含む。除外は対象外条件への該当を書く）",
      "followups": "本文で確認すべき事項",
      "current_status": "提供中 | 提供終了 | 未確認",
      "period_confirmed": true,
      "sources": [{"url": "", "date": "YYYY-MM-DD", "fact": "このURLが裏付ける事実", "is_2026": true}]
    }
  ]
}
```

## 禁止
- スニペットにない数値・日付・URLを作らない。
- 未確認を「実績なし」「非該当」に置き換えない。
- 企業全体・セグメント全体の業績を個別事業の実績にしない。
