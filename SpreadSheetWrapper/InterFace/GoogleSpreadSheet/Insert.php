<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\ValueRange;
use Google\Service\Sheets\AppendValuesResponse;
use Google\Service\Sheets\UpdateValuesResponse;
use SpreadSheetWrapper\Util\EnvLoader;

/**
 * =========================================================
 * Insert
 * ---------------------------------------------------------
 * Googleスプレッドシートへ「追記（append）」と「上書き（update）」を行う
 * ユーティリティクラス。値のネスト正規化も内包します。
 *
 * 【前提】
 * - 認証は SpreadSheetAuth が行い、Client を受け取ります
 * - .env（例：.env.local）に以下を定義
 *     SHEET_ID="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
 *     SHEET_NAME="シート1"
 *
 * 【想定する値の形】
 * - 2次元配列（複数行）
 *   [[A1, B1, C1], [A2, B2, C2]]
 * - 1次元配列（1行のみ）
 *   [A1, B1, C1]  → 内部で [[A1, B1, C1]] に正規化
 * - 内部要素が配列の場合は改行結合 + JSON で安全に文字列化
 *
 * 【主なメソッド】
 * - insert($values, $range, $sheetName): 追記（append）
 * - update($values, $range, $sheetName): 上書き（update）
 *
 * @package SpreadSheetWrapper\InterFace\GoogleSpreadSheet
 * =========================================================
 */
class Insert
{
    /** @var Client 認証済み Google クライアント */
    private Client $client;

    /** @var string スプレッドシートID */
    private string $sheetID;

    /** @var string シート名（タブ名） */
    private string $sheetName;

    /** @var Sheets Sheets サービス */
    private Sheets $sheet;

    /** @var string .env のキー名：スプレッドシートID */
    private const string DEFAULT_SHEET_ID_ENV_KEY = 'SHEET_ID';

    /** @var string .env のキー名：デフォルトのシート名（タブ名） */
    private const string DEFAULT_SHEET_NAME_ENV_KEY = 'SHEET_NAME';

    /**
     * @param string|null           $sheetID     シートID（未指定なら .env の SHEET_ID）
     * @param string|null           $sheetName   シート名（未指定なら .env の SHEET_NAME）
     * @param SpreadSheetAuth|null  $auth        既存の認証オブジェクト（任意）
     * @param string|null           $envFileName 環境ファイル名（例：.env.local）
     * @param string|null           $envFilePath 環境ファイルのディレクトリ
     * @throws Exception
     */
    public function __construct(
        ?string $sheetID = null,
        ?string $sheetName = null,
        ?SpreadSheetAuth $auth = null,
        ?string $envFileName = null,
        ?string $envFilePath = null
    ) {
        // 認証（既存がなければ作る）
        $auth ??= new SpreadSheetAuth($envFileName, $envFilePath);
        $this->client = $auth->getClient();

        // ID とシート名は .env の値がデフォルト
        $sheetID   ??= EnvLoader::getEnv(self::DEFAULT_SHEET_ID_ENV_KEY, $envFilePath, $envFileName);
        $sheetName ??= EnvLoader::getEnv(self::DEFAULT_SHEET_NAME_ENV_KEY, $envFilePath, $envFileName);

        $this->sheetID   = $sheetID;
        $this->sheetName = $sheetName;

        try {
            $this->sheet = new Sheets($this->client);
        } catch (Exception $e) {
            throw new Exception(
                "GoogleSpreadSheet の書き込み準備に失敗しました。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * 値を追記（末尾に追加）します。A1記法の範囲を指定します。
     * 例）$range = "A1" / "A:C" / "A1:C1" など
     *
     * @param array       $values    1行または複数行の配列
     * @param string      $range     追記先の基準レンジ（A1記法）
     * @param string|null $sheetName 対象シート名（未指定なら既定）
     * @return AppendValuesResponse
     * @throws Exception
     */
    public function insert(array $values, string $range, ?string $sheetName = null): AppendValuesResponse
    {
        $sheetName  ??= $this->sheetName;
        $writeRange  = "{$sheetName}!{$range}";
        $valueRange  = $this->buildValueRange($values);

        try {
            /** @var AppendValuesResponse $res */
            $res = $this->sheet->spreadsheets_values->append(
                $this->sheetID,
                $writeRange,
                $valueRange,
                ['valueInputOption' => 'USER_ENTERED']
            );
            return $res;
        } catch (Exception $e) {
            throw new Exception(
                "append に失敗しました（range: {$writeRange}）。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * 値を上書き（指定範囲に更新）します。A1記法の範囲を指定します。
     * 例）$range = "A1:C3" → 3行×3列を上書き
     *
     * @param array       $values    1行または複数行の配列
     * @param string      $range     上書き先レンジ（A1記法）
     * @param string|null $sheetName 対象シート名（未指定なら既定）
     * @return UpdateValuesResponse
     * @throws Exception
     */
    public function update(array $values, string $range, ?string $sheetName = null): UpdateValuesResponse
    {
        $sheetName  ??= $this->sheetName;
        $writeRange  = "{$sheetName}!{$range}";
        $valueRange  = $this->buildValueRange($values);

        try {
            /** @var UpdateValuesResponse $res */
            $res = $this->sheet->spreadsheets_values->update(
                $this->sheetID,
                $writeRange,
                $valueRange,
                ['valueInputOption' => 'USER_ENTERED']
            );
            return $res;
        } catch (Exception $e) {
            throw new Exception(
                "update に失敗しました（range: {$writeRange}）。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * 値配列を ValueRange に変換（内部でネスト正規化）
     *
     * @param array $values
     * @return ValueRange
     */
    private function buildValueRange(array $values): ValueRange
    {
        $this->normalizeNests($values);
        return new ValueRange(['values' => $values]);
    }

    /**
     * ネスト正規化：
     * - 最上位が1次元なら [[...]] に包む
     * - 第2階層の要素に配列が来たら、改行で連結しつつ配列要素は JSON 文字列化
     * - 各行は array_values でキー詰め
     *
     * @param array $values (参照で破壊的に更新)
     * @return void
     */
    private function normalizeNests(array &$values): void
    {
        // 1次元だけ渡されたら行として包む
        $isAssocOrFlat = true;
        foreach ($values as $v) {
            if (is_array($v)) { $isAssocOrFlat = false; break; }
        }
        if ($isAssocOrFlat) {
            $values = [$values];
        }

        foreach ($values as &$row) {
            if (!is_array($row)) {
                $row = [$row];
            }
            foreach ($row as &$cell) {
                if (is_array($cell)) {
                    // 配列セルは安全に文字列化（深い配列は JSON、一次元は改行結合）
                    $cell = implode("\n", array_map(
                        static function ($v) {
                            return is_array($v)
                                ? json_encode($v, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES)
                                : (string)$v;
                        },
                        $cell
                    ));
                }
            }
            unset($cell);
            // 数値添字に詰める
            $row = array_values($row);
        }
        unset($row);
    }
}
