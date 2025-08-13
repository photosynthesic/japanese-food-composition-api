# japanese-food-composition-api

文部科学省の「日本食品標準成分表（八訂）増補2023年」をJSON形式で配信する静的API

## データソース

- [日本食品標準成分表（八訂）増補2023年](https://www.mext.go.jp/a_menu/syokuhinseibun/mext_00001.html)
- 2,538食品の栄養成分データ
- 文部科学省 科学技術・学術審議会 資源調査分科会報告書

## API エンドポイント

- `GET /data/foods.json` - 全食品データ
- `GET /data/categories.json` - 食品分類一覧  
- `GET /data/metadata.json` - メタデータ（バージョン・更新日）

## データ形式

### データの基本仕様

すべての栄養成分データは **可食部100g当たり** で統一されています。

**可食部とは:**
- 実際に食べられる部分のみ（皮、骨、殻などを除く）
- 廃棄率と組み合わせて実際の栄養計算が可能

**例：**
```javascript
{
  "食品名": "鶏卵（全卵）",
  "廃棄率": 12,           // 12%（殻）
  "エネルギーkcal": 142,   // 可食部（卵白+卵黄）100g当たり
  "たんぱく質": 12.2      // 実際の食べられる部分での値
}
```

### 利用時の注意点

1. **実際の摂取量計算**
   ```javascript
   // 鶏卵1個（60g）を食べた場合
   const 実際の摂取カロリー = 142 × 0.6 × (1 - 0.12) = 約75kcal
   ```

2. **廃棄率の考慮**
   - 買い物時：廃棄率を考慮した購入量の計算
   - 調理時：実際の可食部量の把握

3. **データの精度**
   - 天然食品は個体差や季節変動があります
   - 加工食品は製造者により成分が異なる場合があります

### 特殊値の扱い

栄養成分データには以下の特殊値が含まれます：

| 値 | 元データ | 意味 |
|---|---|---|
| `"trace"` | `Tr` | 微量（分析方法では正確に定量できないほど少量だが存在） |
| `"negligible"` | `(0)` `(数値)` | 検出されない/栄養学的に無視できる量 |
| `null` | `-` | 未測定・データなし |
| `"*"` | `*` | 計算値/推定値（※推測）|

**※注記**: `"*"`の正確な意味は公式文書で確認できませんでしたが、主に「利用可能炭水化物」項目で使用されており、他の成分値から計算で求められた推定値と推測されます。

### 使用例

```javascript
// 数値計算時の処理例
function getNutrientValue(value) {
  if (value === "trace" || value === "negligible") return 0;
  if (value === null) return null; // または適切なデフォルト値
  if (value === "*") return null; // 計算値は用途に応じて処理
  return typeof value === "number" ? value : parseFloat(value) || 0;
}
```

## 開発

```bash
# 依存関係のインストール
yarn install

# Excelファイルの構造分析
yarn analyze

# Excel→JSON変換
yarn convert

# データ更新（自動化スクリプト）
yarn update <Excelファイル名>
```

### データ更新の自動化

新しい食品成分表が公開された際は、自動更新スクリプトを使用できます：

```bash
# 例：2024年版への更新
yarn update 20240428-mxt_kagsei-mext_00001_013.xlsx
```

スクリプトは以下を自動実行します：
- Excelファイルの検証
- 変換スクリプトのパス更新
- データ構造の分析
- Excel→JSON変換
- データ品質チェック
- メタデータ更新
- 配信用ファイル生成
- 更新レポート作成

## ライセンス

このプロジェクトは[政府標準利用規約2.0](https://cio.go.jp/sites/default/files/uploads/documents/opendata_nijiriyou_betten1.pdf)（[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/deed.ja)互換）に基づいて提供されています。

### 出典表記

データを利用する際は、以下のように出典を明記してください：

「日本食品標準成分表（八訂）増補2023年」（文部科学省）より引用