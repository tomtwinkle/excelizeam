# `excelize.Style` の取得・再適用範囲

## 対象バージョンと契約

この表は、PR の統合ツリーで使われる Excelize **v2.10.1** の `excelize.Style` を基準にしています。対応する公開型の全40フィールドを棚卸しし、セルの実効書式を取得して `NewStyle` で再適用できる範囲を定義します。元の `NewStyle` 入力を復元することや、任意の OOXML を無損失で複製することは契約に含めません。

PR ブランチ自身の `go.mod` は v2.6.0 です。v2.6.0 には `Style.Lang` があり、`DecimalPlaces` は `int`、`Fill.Transparency` や Font の色参照フィールドはありません。v2.10.1 では `Style.Lang` がなく、`DecimalPlaces` は `*int` です。v2.6.0 固有の `Lang` 指定は取得・再適用対象外です。実装は両方でビルドできるよう、バージョン固有フィールドを実行時に扱います。gradient の shading 番号もバージョン間で異なるため、v2.10.1 の17 preset と v2.6.0 の6 preset を別の表で解釈します。

フィールド定義の根拠は Excelize の固定タグにある [`xmlStyles.go`](https://github.com/qax-os/excelize/blob/v2.10.1/xmlStyles.go)、[`styles.go`](https://github.com/qax-os/excelize/blob/v2.10.1/styles.go)、[`numfmt.go`](https://github.com/qax-os/excelize/blob/v2.10.1/numfmt.go) です。

## フィールド対応表

| 型 | 公開フィールド | 取得・再適用と検証 |
|---|---|---|
| `Style` | `Border`, `Fill`, `Font`, `Alignment`, `Protection` | 下記の各型に従って実効値を取得し、保存・再読込後に同じ書式となることを確認します。 |
| `Style` | `NumFmt`, `CustomNumFmt` | 組み込み ID は `NumFmt`、ユーザー定義コードは完全な書式文字列を `CustomNumFmt` に保持します。組み込み形式と quoted/escaped 文字を含む独自形式を再適用テストします。 |
| `Style` | `DecimalPlaces`, `NegRed` | v2.10.1 の組み込み形式では Excelize の `GetStyle` が返す値を使います。独自形式では decimal 数や `[Red]` を文字列から推測せず、同じ意味を持つ完全な `CustomNumFmt` を保存します。組み込みの decimal 数は個別に再適用テストします。 |
| `Border` | `Type`, `Color`, `Style` | 6種類の辺と style 0〜13を検証します。色は実効 RGB で保持します。 |
| `Fill` | `Type`, `Pattern`, `Color` | pattern 0〜18と単色を検証します。`Color` は foreground を優先し、なければ background を取得します。 |
| `Fill` | `Shading` | v2.10.1 の全17 gradient preset を、geometry、stop 数、stop 位置、終端色まで確認します。v2.6.0 では同バージョンの6 preset を使います。 |
| `Fill` | `Transparency` | v2.10.1 の公開フィールドですが、Excelize では chart / shape 用でセル書式には保存されません。セル書式の取得・再適用対象外です。 |
| `Font` | `Bold`, `Italic`, `Strike` | boolean 属性を保持します。値のない `<b/>`、`<i/>`、`<strike/>` は OOXML の既定に従って `true` とします。各項目の保存後 XML と値を検証します。 |
| `Font` | `Underline`, `Family`, `Size`, `VertAlign`, `Charset` | 取得後に再適用し、XML 属性と再読込値を確認します。v2.10.1 固有の `VertAlign` / `Charset` は v2.6.0 では型にありません。 |
| `Font` | `Color` | v2.10.1 では OOXML に RGB がある場合に限り、Tint 適用前の RGB を保持します。theme / indexed のみの色では空です。これにより同じ tint を二度適用しません。v2.6.0 は色参照フィールドがないため、解決した見た目の RGB を返します。RGB のみで再適用する見た目も別に検証します。 |
| `Font` | `ColorIndexed`, `ColorTheme`, `ColorTint` | v2.10.1 では元の selector と tint を保持し、同じ theme / palette の workbook で再適用できることを確認します。色の selector と見た目の RGB は別々に検証します。v2.6.0 ではこれらの公開フィールドがありません。 |
| `Alignment` | `Horizontal`, `Vertical`, `Indent`, `RelativeIndent`, `ReadingOrder`, `TextRotation`, `WrapText`, `ShrinkToFit`, `JustifyLastLine` | 9項目を全て設定し、返却値と保存・再読込後の値を検証します。 |
| `Protection` | `Hidden`, `Locked` | 両方の値を保持します。XMLで `locked` が省略された場合は Excel の既定である `true` として扱います。保存・再読込を検証します。 |

## 再適用できない入力

- v2.10.1 の gradient は17 preset に対応します。3 stop の preset 2・5・8・11 は、position が `0, 0.5, 1` で、最初と最後の実効色が同じ場合に限り、`Style.Fill.Color` の2色へ縮約できます。gradient の任意 geometry、任意 stop 配置、異なる3つ目の色は公開 `Style` で表現できません。未対応 gradient は種類だけを返し、色を返さないため、そのまま `NewStyle` へ渡すとエラーになります。
- pattern fill は foreground と background の2色を使えますが、`Style.Fill.Color` は1色だけです。両方あるときは foreground を保持し、background は再適用時に失われます。background のみある場合も API の再適用は foreground として出力します。
- `Fill.Transparency` はセル style XML に保存されず、値を復元できません。
- 未知の pattern ID と border style 名は公開列挙値へ写せないため 0 (`none`) として返します。auto 色、未知の色参照、palette / theme から解決できない色は空色になります。
- `Border` には `outline`、`vertical`、`horizontal` の公開表現がありません。取得対象は `left`、`right`、`top`、`bottom`、`diagonalUp`、`diagonalDown` です。未対応の border style 名は `Style` 0 (`none`) に写ります。
- `Font` の `Color` は auto 色を表現しません。`<color auto="1"/>` は空色とし、再適用では自動色指定そのものを出力できません。黒を示す indexed 0 へ誤変換することはありません。theme / indexed の参照値は v2.10.1 で保持できますが、見た目を RGB 固定にしたい場合は `ColorIndexed` を -1、`ColorTheme` を nil、`ColorTint` を 0 にし、表示色を `Color` に指定します。参照 identity と見た目の固定は異なる選択です。同じ identity と見た目を保つテストでは、再適用先の theme / palette が元と同一であることを確認します。
- `ColorIndexed` の zero 値は「未指定」と「index 0」を区別できません。実装は RGB、theme、auto がある場合は selector なしとし、index 単独の zero は index 0 として扱います。公開 API の型に状態を表す値がないため、明示された index 0 と selector 省略の両方を区別できません。
- `formatCode16` は読み込み時に選択した書式コードを `CustomNumFmt` に保持します。`NewStyle` は通常の `formatCode` を出力するため、拡張属性そのものは再生成しません。decimal や `[Red]` を推測して誤るより、選択した完全なコードを保持します。
- OOXML に複数の書式定義や named style があるときは、セルに適用された実効書式を返します。元の XF の構造や継承元の参照関係は保存しません。

## 検証

`style_roundtrip_test.go` のテストは v2.10.1 の40フィールド棚卸し、Font fieldごとの再適用、全 gradient preset、pattern 0〜18、border style 0〜13、theme/indexed identity と RGB、組み込み decimal 数、明示適用された Font ID 0、空 boolean 属性、疎な border、pattern の2色制約を確認します。gradient や色参照などの焦点項目は、独自 extractor の戻り値だけでなく、保存後の `xl/styles.xml` を XML として読み直して比較します。
