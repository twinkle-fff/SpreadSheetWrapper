<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;
use SpreadSheetWrapper\Util\EnvLoader;

/**
 * Google Sheets の「行グループ（アウトライン）」を作成し、必要に応じて折りたたむクラス。
 *
 * - Sheets GUI の「データ → グループ化 → 行をグループ化」と同等の操作を Sheets API(v4) の batchUpdate で行う。
 * - 代表行（親行）を常に表示し、その直下の明細行だけをグループ化して折りたたむ用途を想定。
 *
 * 注意:
 * - Sheets API の DimensionGroup(range) は sheetName ではなく numeric の sheetId を要求するため、
 *   コンストラクタで sheetName → sheetId を解決する。
 * - startIndex は 0-based、endIndex は exclusive（終端は含まれない）。
 */
class SetGroup
{
    /** Google API Client（認証済み） */
    private Client $client;

    /** Google Sheets Service */
    private Sheets $sheets;

    /** Spreadsheet ID（ドキュメント全体のID） */
    private string $spreadsheetId;

    /** Sheet name（タブ名） */
    private string $sheetName;

    /** Sheet ID（タブの numeric id） */
    private int $sheetId;

    /** @var string .env のキー名：スプレッドシートID */
    private const string DEFAULT_SHEET_ID_ENV_KEY = 'SHEET_ID';

    /** @var string .env のキー名：デフォルトのシート名（タブ名） */
    private const string DEFAULT_SHEET_NAME_ENV_KEY = 'SHEET_NAME';

    /**
     * コンストラクタ。
     *
     * - SpreadSheetAuth から認証済み Client を受け取り Sheets Service を生成する。
     * - sheetID / sheetName が未指定の場合は .env から取得する。
     * - sheetName から sheetId（numeric）を解決する。
     *
     * @param string|null $sheetID      Spreadsheet ID。未指定なら .env(SHEET_ID) を使用
     * @param string|null $sheetName    タブ名。未指定なら .env(SHEET_NAME) を使用
     * @param SpreadSheetAuth|null $auth 認証済み Client 生成用。未指定なら生成する
     * @param string|null $envFileName  EnvLoader 用の .env ファイル名（任意）
     * @param string|null $envFilePath  EnvLoader 用の .env パス（任意）
     *
     * @throws Exception Sheets Service 生成・sheetId 解決に失敗した場合
     */
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

        $this->spreadsheetId = $sheetID;
        $this->sheetName = $sheetName;

        try {
            $this->sheets = new Sheets($this->client);
            $this->sheetId = $this->resolveSheetIdByName($this->spreadsheetId, $this->sheetName);
        } catch (Exception $e) {
            throw new Exception(
                "GoogleSpreadSheet の準備に失敗しました。\n詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * 指定された「親（代表）行」の次行〜最終行を 1つの行グループとして作成し、
     * 必要に応じて折りたたみ（collapsed=true）にする。
     *
     * 例:
     * - presentRowNumber=10（代表行）
     * - lastRowNumber=15（同じ親IDの最終行）
     * → 11〜15 行（明細）をグループ化し、代表行(10)は常に表示する。
     *
     * indexOffset について:
     * - 呼び出し側の都合で行番号に補正が必要な場合に使用する（1-basedで加算）。
     *
     * @param int  $presentRowNumber 代表行の行番号（1-based）
     * @param int  $lastRowNumber    同一ブロックの末尾行番号（1-based）
     * @param int  $indexOffset      行番号補正（1-basedで加算）
     * @param bool $willClose        true の場合、作成したグループを折りたたむ
     *
     * @return void
     *
     * @throws Exception Sheets API 呼び出しで失敗した場合（Google\Service\Exception 等）
     */
    public function closeDimensionGroup(
        int $presentRowNumber,
        int $lastRowNumber,
        int $indexOffset = 0,
        bool $willClose = true
    ): void {
        // 明細範囲（1-based）
        $detailStartRow1 = $presentRowNumber + $indexOffset + 1;
        $detailEndRow1   = $lastRowNumber + $indexOffset;

        // 明細が無ければ何もしない
        if ($detailEndRow1 < $detailStartRow1) {
            return;
        }

        // Sheets API の index（0-based / end exclusive）
        $startIndex0 = $detailStartRow1 - 1;
        $endIndex0   = $detailEndRow1;

        $requests = [];

        // グループ作成
        $requests[] = [
            'addDimensionGroup' => [
                'range' => [
                    'sheetId'    => $this->sheetId,
                    'dimension'  => 'ROWS',
                    'startIndex' => $startIndex0,
                    'endIndex'   => $endIndex0,
                ],
            ],
        ];

        // 折りたたみ（任意）
        if ($willClose) {
            $requests[] = [
                'updateDimensionGroup' => [
                    'dimensionGroup' => [
                        'range' => [
                            'sheetId'    => $this->sheetId,
                            'dimension'  => 'ROWS',
                            'startIndex' => $startIndex0,
                            'endIndex'   => $endIndex0,
                        ],
                        'depth' => 1,
                        'collapsed' => true,
                    ],
                    'fields' => 'collapsed',
                ],
            ];
        }

        $body = new BatchUpdateSpreadsheetRequest([
            'requests' => $requests,
        ]);

        // Sheets API: spreadsheets.batchUpdate
        $this->sheets->spreadsheets->batchUpdate($this->spreadsheetId, $body);
    }

    /**
     * sheetName（タブ名）から sheetId（numeric）を解決する。
     *
     * @param string $spreadsheetId Spreadsheet ID
     * @param string $sheetName     タブ名
     *
     * @return int numeric sheetId
     *
     * @throws Exception 指定したタブ名が見つからない場合
     */
    private function resolveSheetIdByName(string $spreadsheetId, string $sheetName): int
    {
        // fields を絞って軽量化
        $ss = $this->sheets->spreadsheets->get($spreadsheetId, [
            'fields' => 'sheets(properties(sheetId,title))'
        ]);

        foreach ($ss->getSheets() as $sheet) {
            $prop = $sheet->getProperties();
            if ($prop && $prop->getTitle() === $sheetName) {
                return (int)$prop->getSheetId();
            }
        }

        throw new Exception("Sheet name not found: {$sheetName}");
    }
}

if(basename(__FILE__) === basename($_SERVER['SCRIPT_FILENAME'])){
    require_once __DIR__."/../../../vendor/autoload.php";
    $sg = new SetGroup();

    $sg->closeDimensionGroup(0,2,3);

}
