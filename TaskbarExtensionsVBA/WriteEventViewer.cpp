/***************************************************************************************************
 *							    シンプルなイベントビューアー書き込み
 ***************************************************************************************************/



#include <windows.h>
#include <tchar.h>



//***************************************************************************************************
//* 機能　　 ： 引数を基に、イベントビュアーへ書き込む
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・eventID         イベントID
//             ・sourceName      ソース
//             ・description     説明
//             ・eventLevel      イベントレベル(エラー、情報、警告)
//             ・category        タスクのカテゴリ
//             ・alwaysLog       モード
//                  →TRUE           常に書き込み
//                  　FALSE          デバッグビルドのみ書き込み
//***************************************************************************************************
void WriteToEventViewer(
    DWORD eventID,
    const wchar_t* sourceName,
    const wchar_t* description,
    WORD eventLevel,
    WORD category,
    BOOL alwaysLog)
{
#if !defined(_DEBUG)
    if (!alwaysLog) {
        // リリースビルドかつ alwaysLog == FALSE → 書き込み禁止
        return;
    }
#endif

    HANDLE hEventLog = RegisterEventSource(NULL, sourceName);
    if (hEventLog) {
        LPCWSTR messages[1] = { description };
        ReportEvent(
            hEventLog,      // Handle to event log
            eventLevel,     // Event type (e.g., ERROR, WARNING, INFORMATION)
            category,       // Event category
            eventID,        // Event identifier
            NULL,           // No user security identifier
            1,              // One substitution string
            0,              // No binary data
            messages,       // Pointer to strings
            NULL            // No binary data
        );
        DeregisterEventSource(hEventLog);
    }
}
