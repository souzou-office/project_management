<?php
/**
 * ============================================
 * Laravel 業務管理ソフト API テンプレート
 * ============================================
 *
 * GASスプレッドシートから案件概要を取得するためのAPIエンドポイント。
 * このファイルをLaravel側に設置して使用してください。
 *
 * 【設置手順】
 * 1. このコントローラーを app/Http/Controllers/Api/ に配置
 * 2. routes/api.php にルートを追加
 * 3. APIトークン認証を設定
 * 4. GAS側の Script Properties に URL とトークンを設定
 */

// ============================================
// routes/api.php に追加するルート
// ============================================
/*
use App\Http\Controllers\Api\CaseSummaryController;

Route::middleware('auth:sanctum')->group(function () {
    Route::get('/cases/{case}/summary', [CaseSummaryController::class, 'summary']);
});
*/

// ============================================
// app/Http/Controllers/Api/CaseSummaryController.php
// ============================================

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use App\Models\CaseRecord; // モデル名は実際のものに合わせてください
use Illuminate\Http\JsonResponse;

class CaseSummaryController extends Controller
{
    /**
     * 案件概要をJSON形式で返す
     *
     * GET /api/cases/{case}/summary
     *
     * レスポンス例:
     * {
     *   "id": 123,
     *   "status": "進行中",
     *   "registration_type": "所有権移転登記",
     *   "client_name": "山田太郎",
     *   "property": "福岡市中央区天神1-1-1 土地200㎡",
     *   "progress": "書類準備中（3/5完了）",
     *   "deadline": "2026-03-15",
     *   "staff": "古江",
     *   "notes": "売主側の印鑑証明待ち"
     * }
     */
    public function summary($caseId): JsonResponse
    {
        // ※ 実際のモデル名・カラム名に合わせて修正してください
        $case = CaseRecord::findOrFail($caseId);

        return response()->json([
            'id'                => $case->id,
            'status'            => $case->status,           // 例: "進行中", "完了", "保留"
            'registration_type' => $case->registration_type, // 例: "所有権移転登記"
            'client_name'       => $case->client_name,       // 例: "山田太郎"
            'property'          => $case->property_summary,  // 例: "福岡市中央区天神1-1-1"
            'progress'          => $case->progress_text,     // 例: "書類準備中（3/5完了）"
            'deadline'          => $case->deadline?->format('Y-m-d'),
            'staff'             => $case->staff?->name,
            'notes'             => $case->notes,
        ]);
    }
}

// ============================================
// APIトークン認証の設定方法（Laravel Sanctum）
// ============================================
/*
【1】Sanctumインストール（未導入の場合）
    composer require laravel/sanctum
    php artisan vendor:publish --provider="Laravel\Sanctum\SanctumServiceProvider"
    php artisan migrate

【2】トークン発行（tinker等で実行）
    $user = User::find(1); // API用ユーザー
    $token = $user->createToken('gas-integration')->plainTextToken;
    // このトークンをGASの Script Properties に BUSINESS_APP_API_TOKEN として設定

【3】GAS側 Script Properties 設定
    BUSINESS_APP_URL       = https://your-app.example.com
    BUSINESS_APP_API_TOKEN = 上記で発行したトークン

【4】動作確認
    curl -H "Authorization: Bearer {token}" \
         -H "Accept: application/json" \
         https://your-app.example.com/api/cases/1/summary
*/
