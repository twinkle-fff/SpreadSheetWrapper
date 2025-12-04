<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;
use SpreadSheetWrapper\Util\EnvLoader;

class SetColor{
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
     * 指定範囲の文字色 / 背景色を設定する
     *
     * @param string      $range            A1 形式の範囲 (例: "A1", "A1:B3", "2:2", "A:Z")
     * @param string|null $color            文字色 HEX (例: "#ff0000" / "ff0000") null の場合は変更しない
     * @param string|null $backGroundColor  背景色 HEX null の場合は変更しない
     * @param string|null $sheetName        シート名 null の場合は $this->sheetName
     * @throws Exception
     */
    public function setColor(
        string $range,
        ?string $color = null,
        ?string $backGroundColor = null,
        ?string $sheetName = null
    ): void {
        // 何も変更しないなら何もしない
        if ($color === null && $backGroundColor === null) {
            return;
        }

        $sheetName ??= $this->sheetName;

        try {
            $sheetId   = $this->getSheetIdByName($sheetName);
            $gridRange = $this->a1ToGridRange($range, $sheetId);

            $userEnteredFormat = [];
            $fields = [];

            if ($color !== null) {
                $userEnteredFormat['textFormat'] = [
                    'foregroundColor' => $this->hexToColorArray($color),
                ];
                $fields[] = 'userEnteredFormat.textFormat.foregroundColor';
            }

            if ($backGroundColor !== null) {
                $userEnteredFormat['backgroundColor'] = $this->hexToColorArray($backGroundColor);
                $fields[] = 'userEnteredFormat.backgroundColor';
            }

            $requests = [
                'requests' => [
                    [
                        'repeatCell' => [
                            'range' => $gridRange,
                            'cell'  => [
                                'userEnteredFormat' => $userEnteredFormat,
                            ],
                            'fields' => implode(',', $fields),
                        ],
                    ],
                ],
            ];

            $body = new BatchUpdateSpreadsheetRequest($requests);

            $this->sheet->spreadsheets->batchUpdate($this->sheetID, $body);
        } catch (Exception $e) {
            throw new Exception(
                "セルの色設定に失敗しました。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * シート名から sheetId を取得
     *
     * @param string $sheetName
     * @return int
     * @throws Exception
     */
    private function getSheetIdByName(string $sheetName): int
    {
        $spreadsheet = $this->sheet->spreadsheets->get($this->sheetID);
        foreach ($spreadsheet->getSheets() as $sheet) {
            $properties = $sheet->getProperties();
            if ($properties->getTitle() === $sheetName) {
                return (int)$properties->getSheetId();
            }
        }

        throw new Exception("指定されたシート名が見つかりません: {$sheetName}");
    }

    /**
     * HEX (#rrggbb / rrggbb / #rgb / rgb) → [red, green, blue] (0.0〜1.0)
     *
     * @param string $hex
     * @return array{red: float, green: float, blue: float}
     * @throws Exception
     */
    private function hexToColorArray(string $hex): array
    {
        $hex = ltrim(trim($hex), '#');

        if (strlen($hex) === 3) {
            // "f0a" → "ff00aa"
            $hex = "{$hex[0]}{$hex[0]}{$hex[1]}{$hex[1]}{$hex[2]}{$hex[2]}";
        }

        if (strlen($hex) !== 6 || !ctype_xdigit($hex)) {
            throw new Exception("不正な HEX カラーコードです: {$hex}");
        }

        $r = hexdec(substr($hex, 0, 2)) / 255;
        $g = hexdec(substr($hex, 2, 2)) / 255;
        $b = hexdec(substr($hex, 4, 2)) / 255;

        return [
            'red'   => $r,
            'green' => $g,
            'blue'  => $b,
        ];
    }

    /**
     * A1 形式の範囲を GridRange 配列に変換
     *
     * 対応例:
     *  - "A1"       → 1セル
     *  - "A1:B3"    → 矩形
     *  - "2:2"      → 行全体
     *  - "A:Z"      → 列全体
     *
     * @param string $a1Range
     * @param int    $sheetId
     * @return array
     * @throws Exception
     */
    private function a1ToGridRange(string $a1Range, int $sheetId): array
    {
        $a1Range = strtoupper(str_replace(' ', '', $a1Range));
        $parts = explode(':', $a1Range);

        $startRef = $parts[0];
        $endRef   = $parts[1] ?? $parts[0];

        $start = $this->parseA1Ref($startRef);
        $end   = $this->parseA1Ref($endRef);

        $range = [
            'sheetId' => $sheetId,
        ];

        // 行
        if ($start['row'] !== null) {
            $range['startRowIndex'] = $start['row'] - 1;
            $range['endRowIndex']   = ($end['row'] ?? $start['row']);
        }

        // 列
        if ($start['col'] !== null) {
            $range['startColumnIndex'] = $start['col'];
            $range['endColumnIndex']   = ($end['col'] ?? $start['col']) + 1;
        }

        return $range;
    }

    /**
     * A1 の片側参照をパース
     *
     * 例:
     *  - "A1" → ['row' => 1, 'col' => 0]
     *  - "A"  → ['row' => null, 'col' => 0]
     *  - "1"  → ['row' => 1, 'col' => null]
     *
     * @param string $ref
     * @return array{row: int|null, col: int|null}
     * @throws Exception
     */
    private function parseA1Ref(string $ref): array
    {
        // 列 + 行 (例: A1, AB10)
        if (preg_match('/^([A-Z]+)(\d+)$/', $ref, $m)) {
            $colLetters = $m[1];
            $rowNumber  = (int)$m[2];

            return [
                'row' => $rowNumber,
                'col' => $this->columnLettersToIndex($colLetters),
            ];
        }

        // 列のみ (例: A, AB)
        if (preg_match('/^[A-Z]+$/', $ref)) {
            return [
                'row' => null,
                'col' => $this->columnLettersToIndex($ref),
            ];
        }

        // 行のみ (例: 1, 10)
        if (preg_match('/^\d+$/', $ref)) {
            return [
                'row' => (int)$ref,
                'col' => null,
            ];
        }

        throw new Exception("不正な A1 形式の参照です: {$ref}");
    }

    /**
     * 列記号 (A, B, ..., Z, AA, AB, ...) → 0始まりの列インデックス
     *
     * @param string $letters
     * @return int
     */
    private function columnLettersToIndex(string $letters): int
    {
        $letters = strtoupper($letters);
        $len = strlen($letters);
        $index = 0;

        for ($i = 0; $i < $len; $i++) {
            $index *= 26;
            $index += (ord($letters[$i]) - ord('A') + 1);
        }

        return $index - 1; // 0-based にする
    }
}

if(basename(__FILE__) === basename($_SERVER['SCRIPT_FILENAME'])){
    require_once __DIR__."/../../../vendor/autoload.php";
    $sc = new SetColor();
    $sc->setColor("A1:A1","ff0000","0000ff","シート3");
    
}
