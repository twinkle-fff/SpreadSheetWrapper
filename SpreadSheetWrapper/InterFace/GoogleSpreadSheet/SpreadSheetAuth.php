<?php
namespace SpreadSheetWrapper\InterFace\GoogleSpreadSheet;

use Exception;
use Google\Client;
use Google\Service\Sheets;
use SpreadSheetWrapper\Util\EnvLoader;

/**
 * ================================================
 * SpreadSheetAuth クラス
 * --------------------------------
 * GoogleスプレッドシートAPIを利用するための認証クラス。
 * .env に設定された認証ファイル（JSON）を読み込み、
 * Google\Client オブジェクトを初期化します。
 *
 * このクラスは初学者でも扱いやすいよう、
 * 認証処理を1つのクラスにまとめています。
 *
 * 【使い方】
 * 1. Google Cloud Platform でサービスアカウントを作成し、
 *    JSON鍵ファイルをダウンロードします。
 * 2. プロジェクト直下の `.env.local` に以下のように書きます：
 *
 *    GOOGLE_CREDENTIAL_PATH_ENV_KEY="credentials/your_key.json"
 *
 * 3. コード内で以下のように呼び出します：
 *
 *    use SpreadSheetWrapper\InterFace\GoogleSpreadSheet\SpreadSheetAuth;
 *
 *    $auth = new SpreadSheetAuth();
 *    $client = $auth->getClient();
 *
 * 4. 取得した $client を使って Sheets API にアクセスします。
 * --------------------------------
 * @package SpreadSheetWrapper\InterFace\GoogleSpreadSheet
 * ================================================
 */
class SpreadSheetAuth
{
    /**
     * .env に記載されたGoogle認証ファイルパスのキー名
     * 例：GOOGLE_CREDENTIAL_PATH_ENV_KEY="credentials/key.json"
     */
    private const GOOGLE_CREDENTIAL_PATH_ENV_KEY = "GOOGLE_CREDENTIAL_PATH_ENV_KEY";

    /** @var Client Google API Client インスタンス */
    private Client $client;

    /**
     * SpreadSheetAuth コンストラクタ
     *
     * @param string|null $envFileName 読み込む.envファイル名（省略可）
     * @param string|null $envFilePath 読み込む.envファイルのディレクトリ（省略可）
     *
     * @throws Exception 認証情報の読み込み・初期化に失敗した場合
     */
    public function __construct(
        ?string $envFileName = null,
        ?string $envFilePath = null
    ) {
        try {
            // Google APIクライアント生成
            $client = new Client();
            $client->setScopes([
                Sheets::SPREADSHEETS,
                Sheets::DRIVE
            ]);

            // 環境変数から認証ファイルのパスを取得
            $credential = EnvLoader::getEnv(
                self::GOOGLE_CREDENTIAL_PATH_ENV_KEY,
                $envFilePath,
                $envFileName
            );

            // 認証設定
            $client->setAuthConfig($credential);

            // インスタンス変数に保持
            $this->client = $client;
        } catch (Exception $e) {
            throw new Exception(
                "Googleスプレッドシートの認証に失敗しました。\n" .
                "詳細: {$e->getMessage()}",
                $e->getCode(),
                $e
            );
        }
    }

    /**
     * Google API Client インスタンスを取得
     *
     * @return Client 初期化済みのGoogle\Clientオブジェクト
     */
    public function getClient(): Client
    {
        return $this->client;
    }
}
