/* global Office */

Office.onReady(() => {
    // コマンドが利用可能になったことを示す
    console.log('Commands are ready');
});

// Office コマンドのハンドラーをここに追加できます
// 例：リボンボタンのクリック時に実行される関数

/**
 * ボタンクリック時にタスクパネルを表示する関数
 */
function showTaskPane(event) {
    // タスクパネンはmanifest.xmlで定義されているため、
    // 通常はOfficeが自動的に処理します。

    // ここに追加のロジックがあれば実装
    console.log('Task pane button clicked');

    // イベント完了を通知（必須）
    event.completed();
}

// グローバル関数として登録（manifest.xmlから参照される場合）
if (typeof global !== 'undefined') {
    global.showTaskPane = showTaskPane;
}