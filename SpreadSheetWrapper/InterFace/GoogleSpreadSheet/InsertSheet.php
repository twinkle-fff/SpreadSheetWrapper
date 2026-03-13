<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use SpreadSheetWrapper\Util\EnvLoader;

class InsertSheet
{
    private Client $client;
    private Sheets $sheet;
    private string $sheetID;
    private string $sheetName;

    private const string DEFAULT_SHEET_ID_ENV_KEY = "SHEET_ID";
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

            $sheetID ??= EnvLoader::getEnv(self::DEFAULT_SHEET_ID_ENV_KEY, $envFilePath, $envFileName);
            $sheetName ??= EnvLoader::getEnv(self::DEFAULT_SHEET_NAME_ENV_KEY, $envFilePath, $envFileName);

            $this->sheetID = $sheetID;
            $this->sheetName = $sheetName;
        } catch (Exception $e) {
            throw new Exception("スプレッドシートの読み込み準備に失敗しました。\n詳細: {$e->getMessage()}", $e->getCode(), $e);
        }
    }

    /**
     * 新しいシートを追加する
     *
     * @param string $sheetName
     * @return void
     * @throws Exception
     */
    public function insertSheet(string $sheetName): void
    {
        try {

            $addSheetRequest = new \Google\Service\Sheets\AddSheetRequest([
                "properties" => [
                    "title" => $sheetName
                ]
            ]);

            $request = new \Google\Service\Sheets\Request([
                "addSheet" => $addSheetRequest
            ]);

            $batchUpdateRequest = new \Google\Service\Sheets\BatchUpdateSpreadsheetRequest([
                "requests" => [$request]
            ]);

            $this->sheet->spreadsheets->batchUpdate(
                $this->sheetID,
                $batchUpdateRequest
            );

        } catch (Exception $e) {
            throw new Exception(
                "シートの追加に失敗しました。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }
}
