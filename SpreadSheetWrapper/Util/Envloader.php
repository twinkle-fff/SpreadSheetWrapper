<?php
namespace SpreadSheetWrapper\Util;

use Dotenv\Dotenv;
use Exception;

/**
 * ================================================
 * EnvLoader クラス
 * --------------------------------
 * .env ファイルを読み込み、指定したキーの値を取得するためのユーティリティ。
 * PHP初心者でも簡単に環境変数を扱えるように設計されています。
 *
 * 【使い方】
 * 1. プロジェクト直下に `.env.local` を作成
 * 2. その中に以下のように設定を書く：
 *
 *    API_KEY=abcdef12345
 *    SHEET_ID=xxxxx
 *
 * 3. PHPコード内で以下のように呼び出す：
 *
 *    use SpreadSheetWrapper\Util\EnvLoader;
 *
 *    $apiKey = EnvLoader::getEnv("API_KEY");
 *
 * --------------------------------
 * ※ このクラスは PHP-Dotenv ライブラリを使用しています。
 *    composer require vlucas/phpdotenv
 * ================================================
 */
class EnvLoader
{
    /** @var string デフォルトの.envファイルの場所（実行中ディレクトリ） */
    private const DEFAULT_ENV_PATH = __DIR__;

    /** @var string デフォルトの.envファイル名 */
    private const DEFAULT_ENV_FILE_NAME = ".env.local";

    /** @var ?Dotenv 現在読み込まれているDotenvインスタンス */
    private static ?Dotenv $dotenv = null;

    /** @var ?string 現在読み込んだ.envファイルのパス */
    private static ?string $envPath = null;

    /** @var ?string 現在読み込んだ.envファイル名 */
    private static ?string $envFileName = null;

    /**
     * 指定したキーの環境変数を取得する
     *
     * @param string $key 取得したい環境変数のキー名
     * @param string|null $envPath 環境ファイルのディレクトリパス（省略可）
     * @param string|null $envFileName 環境ファイル名（省略可）
     *
     * @throws Exception 指定されたキーが存在しない場合
     * @return string 取得した値
     */
    public static function getEnv(
        string $key,
        ?string $envPath = null,
        ?string $envFileName = null
    ): string {
        $envPath ??= self::DEFAULT_ENV_PATH;
        $envFileName ??= self::DEFAULT_ENV_FILE_NAME;

        // .envファイルを読み込む
        self::loadEnv($envPath, $envFileName);

        // 値を取得
        $value = $_SERVER[$key] ?? ($_ENV[$key] ?? null);
        if ($value === null) {
            throw new Exception("環境設定ファイルの読み込みに失敗しました。ファイルに {$key} が設定されていません。");
        }

        return $value;
    }

    /**
     * .envファイルを読み込み、Dotenvインスタンスを初期化
     *
     * @param string $envPath ディレクトリパス
     * @param string $envFileName ファイル名
     *
     * @throws Exception ファイルが見つからない・読み込み失敗時
     * @return Dotenv 読み込まれたDotenvインスタンス
     */
    private static function loadEnv(
        string $envPath,
        string $envFileName
    ): Dotenv {
        // すでに同じファイルが読み込まれている場合は再読み込みしない
        if (
            self::$envPath === $envPath &&
            self::$envFileName === $envFileName &&
            self::$dotenv !== null
        ) {
            return self::$dotenv;
        }

        try {
            self::$envPath = $envPath;
            self::$envFileName = $envFileName;

            $dotenv = Dotenv::createImmutable($envPath, $envFileName);
            $dotenv->load();

            self::$dotenv = $dotenv;
        } catch (Exception $e) {
            throw new Exception(
                "環境設定ファイルの読み込みに失敗しました。" .
                "{$envPath}/{$envFileName} に環境設定ファイルが見つかりません。"
            );
        }

        return self::$dotenv;
    }
}
