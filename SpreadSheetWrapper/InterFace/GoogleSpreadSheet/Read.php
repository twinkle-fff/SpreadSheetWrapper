<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use SpreadSheetWrapper\Util\EnvLoader;

/**
 * ============================================================
 * Read
 * ------------------------------------------------------------
 * Googleスプレッドシートからデータを読み取るクラス。
 *
 * 【主な機能】
 * - 指定範囲を読み取り、列名(A,B,C,...)付きの連想配列として返す。
 * - .env から SHEET_ID / SHEET_NAME を自動読み込み。
 *
 * 【返り値形式】
 * [
 *   ["A" => 値1, "B" => 値2, "C" => 値3],
 *   ["A" => 値4, "B" => 値5, "C" => 値6],
 *   ...
 * ]
 *
 * 【例】
 * $reader = new Read();
 * $rows = $reader->read("A1:C5");
 * ============================================================
 */
class Read
{
    private Client $client;
    private Sheets $sheet;
    private string $sheetID;
    private string $sheetName;

    private const string DEFAULT_SHEET_ID_ENV_KEY   = "SHEET_ID";
    private const string DEFAULT_SHEET_NAME_ENV_KEY = "SHEET_NAME";

    /**
     * @param string|null           $sheetID     シートID（省略時は .env）
     * @param string|null           $sheetName   シート名（省略時は .env）
     * @param SpreadSheetAuth|null  $auth        認証オブジェクト（省略時は自動生成）
     * @param string|null           $envFileName 環境ファイル名
     * @param string|null           $envFilePath 環境ファイルパス
     * @throws Exception
     */
    public function __construct(
        ?string $sheetID = null,
        ?string $sheetName = null,
        ?SpreadSheetAuth $auth = null,
        ?string $envFileName = null,
        ?string $envFilePath = null
    ) {
        try {
            $auth ??= new SpreadSheetAuth($envFileName, $envFilePath);
            $this->client = $auth->getClient();
            $this->sheet = new Sheets($this->client);

            $sheetID   ??= EnvLoader::getEnv(self::DEFAULT_SHEET_ID_ENV_KEY, $envFilePath, $envFileName);
            $sheetName ??= EnvLoader::getEnv(self::DEFAULT_SHEET_NAME_ENV_KEY, $envFilePath, $envFileName);

            $this->sheetID = $sheetID;
            $this->sheetName = $sheetName;
        } catch (Exception $e) {
            throw new Exception("スプレッドシートの読み込み準備に失敗しました。\n詳細: {$e->getMessage()}", $e->getCode(), $e);
        }
    }

    /**
     * 指定範囲を読み取り、列名(A,B,C,...)をキーにした配列を返す
     *
     * @param string      $range     例："A1:C10"
     * @param string|null $sheetName シート名（未指定ならデフォルト）
     * @return array<array<string,mixed>> 読み取ったデータ配列
     * @throws Exception
     */
    public function read(string $range, ?string $sheetName = null): array
    {
        $sheetName ??= $this->sheetName;
        $readRange = "{$sheetName}!{$range}";

        try {
            $response = $this->sheet->spreadsheets_values->get($this->sheetID, $readRange);
            $values = $response->getValues() ?? [];

            if (empty($values)) {
                return [];
            }

            return $this->attachColumnLabels($values);
        } catch (Exception $e) {
            throw new Exception("スプレッドシートの読み込みに失敗しました。\n詳細: {$e->getMessage()}", $e->getCode(), $e);
        }
    }

    /**
     * 列番号から列名(A,B,C,...)を生成して配列に変換
     *
     * @param array $values スプレッドシートから取得した2次元配列
     * @return array<array<string,mixed>>
     */
    private function attachColumnLabels(array $values): array
    {
        $result = [];
        foreach ($values as $rowIndex => $row) {
            $assoc = [];
            foreach ($row as $colIndex => $cell) {
                $colLabel = $this->columnIndexToLetter($colIndex);
                $assoc[$colLabel] = $cell;
            }
            $result[] = $assoc;
        }
        return $result;
    }

    /**
     * 0始まりの列インデックスをExcel風列名(A,B,...,AA,AB,...)に変換
     *
     * @param int $index
     * @return string
     */
    private function columnIndexToLetter(int $index): string
    {
        $letter = '';
        $index += 1;
        while ($index > 0) {
            $mod = ($index - 1) % 26;
            $letter = chr(65 + $mod) . $letter;
            $index = intval(($index - $mod) / 26);
        }
        return $letter;
    }
}
