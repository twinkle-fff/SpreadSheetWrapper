<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;
use Google\Service\Sheets\DeleteDimensionRequest;
use Google\Service\Sheets\DeleteSheetRequest;
use Google\Service\Sheets\DimensionRange;
use Google\Service\Sheets\Request;
use SpreadSheetWrapper\Util\EnvLoader;

/**
 * =========================================================
 * Delete
 * ---------------------------------------------------------
 * Googleスプレッドシートの「行削除」と「シート削除」を行う
 * ユーティリティクラス。
 *
 * 【前提】
 * - 認証は SpreadSheetAuth が行い、Client を受け取ります
 * - .env（例：.env.local）に以下を定義
 *     SHEET_ID="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
 *     SHEET_NAME="シート1"
 *
 * 【主なメソッド】
 * - deleteRow($range, $sheetName): 指定範囲に含まれる行を削除
 * - deleteSheet($sheetName): 指定シートを削除
 *
 * 【注意】
 * - deleteSheet は誤削除防止のため sheetName を必須とします
 * - deleteRow の range は A2:Y2 / A2 / 2:2 のような行番号を含む形式を想定します
 * - A2:Y のように終了行がない場合は、開始行のみ削除します
 *
 * @package SpreadSheetWrapper\InterFace\GoogleSpreadSheet
 * =========================================================
 */
class Delete
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

    public function __construct(
        ?string $sheetID = null,
        ?string $sheetName = null,
        ?SpreadSheetAuth $auth = null,
        ?string $envFileName = null,
        ?string $envFilePath = null
    ) {
        $auth ??= new SpreadSheetAuth($envFileName, $envFilePath);
        $this->client = $auth->getClient();

        $sheetID   ??= EnvLoader::getEnv(self::DEFAULT_SHEET_ID_ENV_KEY, $envFilePath, $envFileName);
        $sheetName ??= EnvLoader::getEnv(self::DEFAULT_SHEET_NAME_ENV_KEY, $envFilePath, $envFileName);

        $this->sheetID = $sheetID;
        $this->sheetName = $sheetName;

        try {
            $this->sheet = new Sheets($this->client);
        } catch (Exception $e) {
            throw new Exception(
                "GoogleSpreadSheet の削除準備に失敗しました。\n詳細: {$e->getMessage()}",
                (int)$e->getCode(),
                $e
            );
        }
    }

    /**
     * 指定範囲に含まれる行を削除する。
     *
     * @param string $range 削除対象行を含む範囲（例: A2:Y2 / A2 / 2:2）
     * @param string|null $sheetName 対象シート名。省略時は既定シート名
     * @throws Exception
     * @return bool
     */
    public function deleteRow(string $range, ?string $sheetName = null): bool
    {
        $sheetName ??= $this->sheetName;

        $sheetId = $this->resolveSheetId($sheetName);
        [$startIndex, $endIndex] = $this->parseRowIndexes($range);

        $request = new Request([
            'deleteDimension' => new DeleteDimensionRequest([
                'range' => new DimensionRange([
                    'sheetId' => $sheetId,
                    'dimension' => 'ROWS',
                    'startIndex' => $startIndex,
                    'endIndex' => $endIndex,
                ]),
            ]),
        ]);

        $body = new BatchUpdateSpreadsheetRequest([
            'requests' => [$request],
        ]);

        try {
            $this->sheet->spreadsheets->batchUpdate($this->sheetID, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception(
                "行の削除に失敗しました。sheetName={$sheetName}, range={$range}\n詳細: {$e->getMessage()}",
                (int)$e->getCode(),
                $e
            );
        }
    }

    /**
     * 指定シートを削除する。
     *
     * 誤削除防止のため、sheetName は必須。
     *
     * @param string $sheetName 削除対象シート名
     * @throws Exception
     * @return bool
     */
    public function deleteSheet(string $sheetName): bool
    {
        if (trim($sheetName) === '') {
            throw new Exception('deleteSheet の sheetName は必須です。');
        }

        $sheetId = $this->resolveSheetId($sheetName);

        $request = new Request([
            'deleteSheet' => new DeleteSheetRequest([
                'sheetId' => $sheetId,
            ]),
        ]);

        $body = new BatchUpdateSpreadsheetRequest([
            'requests' => [$request],
        ]);

        try {
            $this->sheet->spreadsheets->batchUpdate($this->sheetID, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception(
                "シートの削除に失敗しました。sheetName={$sheetName}\n詳細: {$e->getMessage()}",
                (int)$e->getCode(),
                $e
            );
        }
    }

    /**
     * シート名から sheetId を取得する。
     *
     * @param string $sheetName
     * @throws Exception
     * @return int
     */
    private function resolveSheetId(string $sheetName): int
    {
        $spreadsheet = $this->sheet->spreadsheets->get($this->sheetID);

        foreach ($spreadsheet->getSheets() as $sheet) {
            $properties = $sheet->getProperties();

            if ($properties->getTitle() === $sheetName) {
                return (int)$properties->getSheetId();
            }
        }

        throw new Exception("指定されたシートが見つかりません。sheetName={$sheetName}");
    }

    /**
     * A1形式の範囲から削除対象行の startIndex / endIndex を取得する。
     *
     * Google Sheets API の行番号は0始まり、endIndexは排他的。
     *
     * @param string $range
     * @throws Exception
     * @return array{0:int,1:int}
     */
    private function parseRowIndexes(string $range): array
    {
        if (preg_match('/^(\d+):(\d+)$/', $range, $m)) {
            $startRow = (int)$m[1];
            $endRow = (int)$m[2];

            return [$startRow - 1, $endRow];
        }

        if (preg_match('/^[A-Z]+(\d+)(?::[A-Z]+(\d+))?$/i', $range, $m)) {
            $startRow = (int)$m[1];
            $endRow = isset($m[2]) && $m[2] !== ''
                ? (int)$m[2]
                : $startRow;

            return [$startRow - 1, $endRow];
        }

        throw new Exception("削除対象行を解釈できません。range={$range}");
    }
}
