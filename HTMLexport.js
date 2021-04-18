// HTMLexport.js Ver.1.6.0 (c) 2021 Pyo (PapparaSoft)

// 出力設定
var option_indexmode            = 0;                    // 「index.html」出力モード（-1:出力しない 0:見出し一覧 1:最新日記の月）
var option_reverse              = false;                // 日記の並びを逆順にする
var option_exportyeardiary      = true;                 // １年間の日記を出力（「2014.html」「2015.html」など）
var option_exportalldiary       = false;                // すべての日記を出力（「all.html」）
var option_sidebar              = true;                 // サイドバー表示（カレンダーなど）
var option_calendarstartweek    = 0;                    // サイドバーカレンダーの開始曜日（0:日 1:月 2:火 3:水 4:木 5:金 6:土）
var option_navigationtop        = true;                 // 上部にナビゲーションリンクを表示
var option_navigationbottom     = true;                 // 下部にナビゲーションリンクを表示
var option_showdateinfo         = true;                 // 日記の日付の部分に曜日と祝日の名前を表示
// 変換設定
var option_diarytitle           = true;                 // 日記の一行目を自動的に見出しに変換
var option_enablehtmltag        = false;                // 日記本文のHTMLタグを有効にする（※表示が崩れることがあるためあまりおすすめしません。）
var option_ptagconvert          = true;                 // <p>タグ変換（falseにすると、改行を単純に<br>に変換します。）
var option_indent               = 0;                    // 変換後のHTMLのインデント（-1:なし 0:タブ 1～:スペース個数）
var option_charcode             = 'utf-8';              // 変換後のHTMLの文字コード（'utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp'など）
var option_linecode             = '\r\n';               // 変換後のHTMLの改行コード（CRLF:'\r\n' LF:'\n' CR:'\r'）
var option_closetagslash        = false;                // 閉じスラッシュを付けるかどうか（<br>→<br />、<img ～>→<img ～ />、など）
// 印刷設定
var option_printpagebreak       = false;                // 日記の始まりで常に改ページ印刷（CCSのpage-break-beforeプロパティを使用）
// その他
var option_startmessage         = true;                 // 変換開始時のメッセージを表示するかどうか



// -------------------------------------------------------------------------------------------------

var g_fs = new ActiveXObject('Scripting.FileSystemObject'); // ファイル操作関連（※ファイルコピーやファイルの列挙に使用）
var g_as = new ActiveXObject('ADODB.Stream');               // データアクセス（※文字コードを扱うファイルの読み書きに使用）
var g_list_filename = [];                                   // ファイル名一覧（「201401.txt」「201402.txt」など）
var g_hash_diarytext = {};                                  // すべての日記を格納する連想配列（キーは'20140101'など）
var g_hash_nav = {};                                        // ナビゲーションリンク用（月別の日記ページへのリンク）
var g_newdiaryyear = 0, g_newdiarymonth = 0;                // 最新日記の年月
var g_now = new Date();                                     // 現在日時（コピーライト用）

main();
WScript.Quit();

// *************************************************************************************************
// メイン関数
// *************************************************************************************************
function main()
{
    // 定義
    var list_week = ['日', '月', '火', '水', '木', '金', '土'];
    var list_holiday = [
        '元日',         '成人の日',     '建国記念の日', '春分の日',     '昭和の日',     '憲法記念日',       'みどりの日',
        'こどもの日',   '海の日',       '山の日',       '敬老の日',     '秋分の日',     '体育の日',         '文化の日',
        '勤労感謝の日', '天皇誕生日',   '振替休日',     '国民の休日',   '新天皇即位',   '即位礼正殿の儀',   'スポーツの日'
    ];
    // 日記フォルダ列挙
    var folderinfo;                                     // フォルダ情報（日記フォルダ列挙用）
    var enum_folder;                                    // フォルダ列挙用
    var foldername = '';                                // フォルダ名（日記の名前）
    var cnt_folder = 0;                                 // 日記フォルダの数
    var cnt_file = 0;                                   // 日記データの数
    // 結果表示用
    var list_out_folder = [];                           // 日記フォルダの名前
    var list_out_file = [];                             // 日記データの数
    var str_message = '';                               // メッセージボックスのテキスト
    var err_cnt = 0;                                    // ファイル保存失敗のカウント
    // 日記本文
    var str_oneday = '';                                // 一日分の日記HTML
    var str_diarymonth = '';                            // 一ヶ月単位の日記HTML
    var str_diaryyear = '';                             // 一年単位の日記HTML
    var str_diaryall = '';                              // すべての日記HTML
    var hash_tmp;                                       // 一時的な連想配列
    // その他
    var year, month, day;                               // 年月日
    var year_old = 0, month_old = 0;                    // 直前の年と月
    var holiday = 0;                                    // 祝日かどうか
    var week = 0;                                       // 曜日（0:日 1:月 2:火 3:水 4:木 5:金 6:土）
    var convert = '';                                   // 日記本文の変換後のHTML
    var key = '';                                       // 連想配列のキー
    var i;                                              // ループ用


    // 変換開始時のメッセージ
    if (option_startmessage) {
        WScript.Echo('HTML変換を行います。\n変換には多少時間がかかります。');
    }

    // スクリプトがあるフォルダのフォルダ情報取得
    folderinfo = g_fs.GetFolder(g_fs.GetParentFolderName(WScript.ScriptFullName));
    // 列挙用のオブジェクトを作成
    enum_folder = new Enumerator(folderinfo.SubFolders);
    // フォルダ内のフォルダをすべて列挙
    for (; !enum_folder.atEnd(); enum_folder.moveNext()) {
        // フルパスからフォルダ名を取り出す
        foldername = g_fs.GetFileName(enum_folder.item());
        // 「_」から始まるフォルダは除く
        if (foldername.match(/^_/)) {
            continue;
        }

        // -----------------------------------------------------------------------------------------
        // すべての日記データを取り出す
        // -----------------------------------------------------------------------------------------
        // 日記ファイル名一覧を取得（日記フォルダの中のファイル名を配列に取り出す）
        g_list_filename = getDiaryFileList(enum_folder.item());
        // 日記ファイル名一覧からナビゲーションリンク用の連想配列を作成
        g_hash_nav = getLinkHash(g_list_filename);
        // 連想配列に取り出す
        for (cnt_file = 0; cnt_file < g_list_filename.length; ++cnt_file) {
            // ファイルを読み込み日記を切り分けて連想配列に代入
            hash_tmp = getDiary(enum_folder.item(), g_list_filename[cnt_file]);
            // 連想配列を結合
            g_hash_diarytext = hashMerge(g_hash_diarytext, hash_tmp);
        }
        // すべての日記データを日付順にソート
        g_hash_diarytext = hashSort(g_hash_diarytext, option_reverse);

        // -----------------------------------------------------------------------------------------
        // HTML出力
        // -----------------------------------------------------------------------------------------
        for (key in g_hash_diarytext) {
            // キーがある場合
            if (g_hash_diarytext.hasOwnProperty(key)) {
                // キーから年月日を取り出して数値化
                year = Number(key.substr(0, 4));    // 年
                month = Number(key.substr(4, 2));   // 月
                day = Number(key.substr(6, 2));     // 日

                // 月が異なる場合、かつ、月の文字列がある場合
                if (month !== month_old && str_diarymonth !== '') {
                    // HTMLファイル保存
                    saveHTMLFile(enum_folder.item(), foldername, year_old, month_old, str_diarymonth, 0);
                    // 初期化
                    str_diarymonth = '';
                }
                // 一年間の日記を出力する場合、かつ、年が異なる場合、かつ、年の文字列がある場合
                if (option_exportyeardiary && year !== year_old && str_diaryyear !== '') {
                    // HTMLファイル保存
                    saveHTMLFile(enum_folder.item(), foldername, year_old, 0, str_diaryyear, 1);
                    // 初期化
                    str_diaryyear = '';
                }

                // 日記本文をHTML変換
                convert = convertHTML(g_hash_diarytext[key]);
                // 行頭にタブを挿入（HTMLの見た目のため）
                convert = convert.replace(/\n/g, '\n\t\t\t\t\t');

                // 曜日を取得
                week = getWeekDay(year, month, day);
                // 指定日が祝日（休日）かどうか
                holiday = getHoliday(year, month, day);
                // 日記のHTMLを作成
                str_oneday = '\t\t\t<article' + (option_printpagebreak ? ' class="pagebreak"' : '') + '>\n';
                str_oneday += '\t\t\t\t<h2 id="diary' + year + formatNumber(month) + formatNumber(day) + '">';
                str_oneday += year + '-' + formatNumber(month) + '-' + formatNumber(day);
                if (option_showdateinfo) {
                    // 曜日
                    if (week === 0 || holiday > 0) {    str_oneday += ' <span class="diary-title-week-sunday">(' + list_week[week] + ')</span>';
                    } else if (week === 6) {            str_oneday += ' <span class="diary-title-week-saturday">(' + list_week[week] + ')</span>';
                    } else {                            str_oneday += ' <span class="diary-title-week">(' + list_week[week] + ')</span>';
                    }
                    // 祝日の場合
                    if (holiday > 0) {
                        str_oneday += ' <span class="diary-title-holiday">' + list_holiday[holiday - 1] + '</span>';
                    }
                }
                str_oneday += '</h2>\n';
                str_oneday += '\t\t\t\t<div class="diary">\n';
                str_oneday += '\t\t\t\t\t' + convert + '\n';
                str_oneday += '\t\t\t\t</div>\n';
                str_oneday += '\t\t\t</article>\n';
                str_oneday += '\n';

                // 日記本文を追加していく
                str_diarymonth += str_oneday;
                if (option_exportyeardiary) str_diaryyear += str_oneday;
                if (option_exportalldiary) str_diaryall += str_oneday;

                year_old = year;
                month_old = month;
            }
        }
        // 月の文字列がある場合
        if (str_diarymonth !== '') {
            // HTMLファイル保存
            if (!saveHTMLFile(enum_folder.item(), foldername, year_old, month_old, str_diarymonth, 0)) {
                err_cnt++;
            }
            // 初期化
            str_diarymonth = '';
        }
        // 一年間の日記を出力する場合、かつ、年の文字列がある場合
        if (option_exportyeardiary && str_diaryyear !== '') {
            // HTMLファイル保存
            if (!saveHTMLFile(enum_folder.item(), foldername, year_old, 0, str_diaryyear, 1)) {
                err_cnt++;
            }
            // 初期化
            str_diaryyear = '';
        }
        // 「all.html」を出力する場合、かつ、文字列がある場合
        if (option_exportalldiary && str_diaryall !== '') {
            // HTMLファイル保存
            if (!saveHTMLFile(enum_folder.item(), foldername, 0, 0, str_diaryall, 2)) {
                err_cnt++;
            }
            // 初期化
            str_diaryall = '';
        }


        // HTMLファイルを１つでも出力した場合
        if (cnt_file > 0) {
            // 目次ページ（index.html）の出力
            if (option_indexmode === 0) {
                if (!createIndexPage(enum_folder.item(), foldername)) {
                    err_cnt++;
                }
            }
            // CSS一式のフォルダがスクリプトファイルと同じディレクトリにある場合
            if (g_fs.FolderExists('_css')) {
                // 「_css」フォルダを「css」の名前でコピー
                g_fs.CopyFolder('_css', enum_folder.item() + '\\css');
            }
        }


        // 結果表示用に日記フォルダの数をカウント
        list_out_folder[cnt_folder] = foldername;
        list_out_file[cnt_folder] = cnt_file;
        cnt_folder++;

        // 初期化
        for (key in g_hash_diarytext) if (g_hash_diarytext.hasOwnProperty(key)) delete g_hash_diarytext[key];
        for (key in g_hash_nav) if (g_hash_nav.hasOwnProperty(key)) delete g_hash_nav[key];
    }


    // ---------------------------------------------------------------------------------------------
    // 結果表示
    // ---------------------------------------------------------------------------------------------
    // 日記フォルダがない場合
    if (cnt_folder <= 0) {
        str_message = '日記フォルダがありません。';
    } else {
        str_message = 'HTML変換が完了しました。\r\n\r\n';
        for(i = 0; i < list_out_folder.length; ++i) {
            str_message += '『' + list_out_folder[i] + '』 … ';
            if (list_out_file[i] > 0) {
                str_message += list_out_file[i] + 'ヶ月分\r\n';
            } else {
                str_message += '日記データが見つかりません\r\n';
            }
        }
        // ファイル保存に失敗している場合
        if (err_cnt > 0) {
            str_message += '\r\n※' + err_cnt + '個のファイル保存に失敗しています。';
        }
    }
    // メッセージ表示
    WScript.Echo(str_message);
}



// *************************************************************************************************
// HTMLファイル保存
// 【引数】
//      path                日記フォルダのパス（例「c:\wDiary\MyDiary」など※最後の\は付けない）
//      foldername          フォルダ名（日記の名前）
//      year                年
//      month               月
//      str_diary           HTML化した日記本文（一ヵ月または一年分）
//      mode                保存モード（0:月単位 1:年単位 2:すべて）
// *************************************************************************************************
function saveHTMLFile(path, foldername, year, month, str_diary, mode)
{
    var str = '';                                       // HTML全体
    var str_title = '';                                 // タイトル
    var html_sidebar = '';                              // サイドバー（カレンダーと月別ページ一覧のリンク）
    var html_navigation = '';                           // 月別のリンク


    // 月単位
    if (mode === 0) {
        // サイドバーHTML作成
        html_sidebar = createSideBarHTML(year, month);
        // ナビゲーションリンクHTML作成
        html_navigation = createNavHTML(year, month);
        // 「index.html」を最新日記の月にする場合、かつ、最新日記の年月の場合
        if (option_indexmode === 1 && (year === g_newdiaryyear && month === g_newdiarymonth)) {
            str_title = foldername;
        } else {
            str_title = year + '年' + month + '月 - ' + foldername;
        }
    // 年単位
    } else if (mode === 1) {
        // ナビゲーションリンクHTML作成
        html_navigation = createNavHTML(year, 0);
        str_title = year + '年 - ' + foldername;
    // すべて
    } else {
        // ナビゲーションリンクHTML作成
        html_navigation = createNavHTML(0, 0);
        str_title = foldername;
    }

    str += '<!DOCTYPE html>\n';
    str += '<html>\n';
    str += '<head>\n';
    str += '\t<meta charset="' + option_charcode + '">\n';
    str += '\t<meta name="viewport" content="initial-scale=1.0">\n';
    str += '\t<title>' + str_title + '</title>\n';
    str += '\t<link rel="stylesheet" href="css/style.css">\n';
    str += '</head>\n';
    str += '<body' + (html_sidebar ? '' : ' class="noside"') + '>\n';
    str += '\t<header>\n';
    str += '\t\t<h1 class="diary-name">' + ((option_indexmode >= 0) ? ('<a href="index.html">' + foldername + '</a>') : foldername) + '</h1>\n';
    str += '\t</header>\n';
    str += '\n';
    str += (option_navigationtop ? (html_navigation + '\n') : '');          // ナビゲーションリンク
    str += '\t<div class="container">\n';
    str += '\t\t<div class="contents">\n';
    str += str_diary;                                                       // 日記本文
    str += '\t\t</div>\n';
    str += '\t</div>\n';
    str += '\n';
    str += (html_sidebar ? (html_sidebar + '\n') : '');                     // サイドバー
    str += (option_navigationbottom ? (html_navigation + '\n') : '');       // ナビゲーションリンク
    str += '\t<footer>\n';
    str += '\t\t<div class="copyright">\n';
    str += '\t\t\t&copy; ' + g_now.getFullYear() + ((option_indexmode >= 0) ? (' <a href="index.html">' + foldername + '</a>') : (' ' + foldername)) + ' All rights reserved.<br>\n';
    str += '\t\t\tPowered by &copy; <a href="http://www.cc9.ne.jp/~pappara/wdiary.html">wDiary</a>\n';
    str += '\t\t</div>\n';
    str += '\t</footer>\n';
    str += '\n';
    str += '</body>\n';
    str += '</html>\n';


    // HTML最終整形
    str = formatHTML(str);

    // ファイル書き込み
    if (mode === 0) {
        // 「index.html」を最新日記の月にする場合、かつ、最新日記の年月の場合
        if (option_indexmode === 1 && (year === g_newdiaryyear && month === g_newdiarymonth)) {
            return saveFile(path + '\\index.html', str);
        } else {
            return saveFile(path + '\\' + year + formatNumber(month) + '.html', str);
        }
    } else if (mode === 1) {
        return saveFile(path + '\\' + year + '.html', str);
    } else {
        return saveFile(path + '\\all.html', str);
    }
}



// *************************************************************************************************
// 日記ファイル一覧を取得
// *************************************************************************************************
function getDiaryFileList(folderpath)
{
    var arr = [];                                       // ファイル名一覧の配列
    var folderinfo;                                     // フォルダ情報
    var enum_file;                                      // ファイル列挙用
    var filename;                                       // ファイル名

    // フォルダ情報取得
    folderinfo = g_fs.GetFolder(folderpath);
    // 列挙用のオブジェクトを作成
    enum_file = new Enumerator(folderinfo.Files);
    // フォルダ内のフォルダをすべて列挙
    for (; !enum_file.atEnd(); enum_file.moveNext()) {
        // フルパスからファイル名を取り出す
        filename = g_fs.GetFileName(enum_file.item());
        // 日記データではない場合（日記データのファイル名は「201401.txt」など）
        if (!filename.match(/^\d{4}\d{2}\.txt$/)) {
            continue;
        }
        // ファイル名を配列に追加
        arr.push(filename);
    }
    // 日記データがない場合
    if (arr.length <= 0) return arr;
    // ソート
    arr.sort();
    if (option_reverse) {
        arr.reverse();
    }
    // 最新日記の年月を取得しておく
    if (option_reverse) {
        g_newdiaryyear = Number(arr[0].substr(0, 4));                   // 年
        g_newdiarymonth = Number(arr[0].substr(4, 2));                  // 月
    } else {
        g_newdiaryyear = Number(arr[arr.length - 1].substr(0, 4));      // 年
        g_newdiarymonth = Number(arr[arr.length - 1].substr(4, 2));     // 月
    }
    // ファイル名一覧の配列を返す
    return arr;
}



// *************************************************************************************************
// ファイルを読み込み日記を切り分けて連想配列に代入
// 【引数】
//      path                日記フォルダのパス（例「c:\wDiary\MyDiary」など※最後の\は付けない）
//      filename            日記データのファイル名（例「201401.txt」など）
// *************************************************************************************************
function getDiary(path, filename)
{
    var ret = {};                                       // 戻り値用の連想配列
    var str_text = '';                                  // ファイル内容
    var str_diary = '';                                 // 日記本文
    var lines = [];                                     // 一行ごと
    var year, month, day = 1;                           // 日
    var key = '';                                       // 連想配列のキー
    var i;                                              // ループ用


    // 日記ファイル名から年月を取り出して数値化
    year = Number(filename.substr(0, 4));       // 年
    month = Number(filename.substr(4, 2));      // 月

    // ファイル読み込み
    str_text = loadFile(path + '\\' + filename);
    // ファイル内容がない場合
    if (str_text.length <= 0) {
        return '';
    }

    // ---------------------------------------------------------------------------------------------
    // 1日単位に分解して連想配列に代入
    // ---------------------------------------------------------------------------------------------
    // 改行を\nに揃える
    str_text = str_text.replace(/\r\n|\r/g, '\n');
    // 改行で区切って配列に代入
    lines = str_text.split('\n');
    // 一行ずつ処理
    for (i = 0; i < lines.length; i++) {
        // 日付の行かどうか
        if (lines[i].match(/^\d{4}\/\d{2}\/\d{2}/)) {
            // これまでに日記があった場合
            if (str_diary !== '') {
                // 連想配列のキーを作成
                key = year + formatNumber(month) + formatNumber(day);
                // 日記本文の連想配列に代入
                ret[key] = str_diary;
                // 日記本文の変数を初期化
                str_diary = '';
            }
            // 日付の日を取り出す（※年月はファイル名の年月を使う）
            day = Number(lines[i].substr(8, 2));
        } else {
            // 日記本文に追加していく
            str_diary += lines[i] + '\n';
        }
    }
    // 日記があった場合
    if (str_diary !== '') {
        // 連想配列のキーを作成
        key = year + formatNumber(month) + formatNumber(day);
        // 日記本文の連想配列に代入
        ret[key] = str_diary;
    }
    // 連想配列を返す
    return ret;
}



// *************************************************************************************************
// 指定日の日記があるかどうか
// *************************************************************************************************
function isDiary(year, month, day)
{
    return g_hash_diarytext.hasOwnProperty(year + formatNumber(month) + formatNumber(day));
}



// *************************************************************************************************
// 日記ファイル名一覧からナビゲーションリンク用の連想配列を作成
// *************************************************************************************************
function getLinkHash(arr)
{
    var ret = {};                                       // 戻り値用の連想配列
    var year, month;                                    // 年と月
    var year_old = 0;                                   // 直前の年
    var i;                                              // ループ用

    for (i = 0; i < arr.length; ++i) {
        // 日記ファイル名から年月を取り出して数値化
        year = Number(arr[i].substr(0, 4));     // 年
        month = Number(arr[i].substr(4, 2));    // 月
        // 年が変わった場合
        if (year !== year_old) {
            ret[year] = ' ' + month + ' ';
            year_old = year;
        } else {
            ret[year] += ' ' + month + ' ';
        }
    }
    // 連想配列を返す
    return ret;
}



// *************************************************************************************************
// サイドバーHTML作成
// *************************************************************************************************
function createSideBarHTML(year, month)
{
    var str = '';                                                   // サイドバーのHTML全体
    // カレンダー
    var list_week = ['日', '月', '火', '水', '木', '金', '土'];     // 曜日の文字
    var calline;                                                    // カレンダーの行数
    var tableval = [];                                              // 表内容の配列
    var monthdays;                                                  // 月の日数
    var startweek;                                                  // 月初めの曜日（0:日 1:月 2:火 3:水 4:木 5:金 6:土）
    var day;                                                        // 日
    var week = 0;                                                   // 曜日
    var holiday = 0;                                                // 祝日かどうか
    var sunday = 0;                                                 // 日曜日かどうか
    var saturday = 0;                                               // 土曜日かどうか
    var str_date = '';                                              // 日付の文字列（「20140101」など）
    var x, y;                                                       // ループ用
    // 日記ページ一覧のリンク
    var year_old = '';                                              // 直前の年の文字列
    var year_tmp, month_tmp;                                        // 一時用
    var num;                                                        // 指定月の日記件数
    var i;                                                          // ループ用


    // サイトバーを表示しない場合
    if (!option_sidebar) {
        return '';
    }

    // ---------------------------------------------------------------------------------------------
    // サイドバーのカレンダー
    // ---------------------------------------------------------------------------------------------
    str += '\t<div class="sidebar">\n';
    str += '\t\t<section>\n';
    str += '\t\t\t<h4 class="caption-side">Calendar</h4>\n';
    str += '\t\t\t<div class="calendar">\n';
    str += '\t\t\t\t<div class="date">' + year + '年' + formatNumber(month) + '月</div>\n';
    str += '\t\t\t\t<table>\n';

    // 月の日数を取得
    monthdays = new Date(year, month, 0).getDate();         // ※月は0～11で指定する必要があるので注意。日を0にすると前月の末日になる
    // 月初めの曜日を取得
    startweek = getWeekDay(year, month, 1);
    // カレンダーの日付部分の行数を求める
    calline = Math.ceil((startweek + monthdays) / 7);

    // 前処理
    week = startweek - option_calendarstartweek;
    sunday = 0 - option_calendarstartweek;
    saturday = 6 - option_calendarstartweek;
    if (week < 0) week += 7;
    if (sunday < 0) sunday += 7;
    if (saturday < 0) saturday += 7;

    // 表内容の配列を初期化
    for (i = 0; i < 7 * calline; i++) tableval[i] = '';
    // 表内容の配列に日付を代入
    for (i = 0; i < monthdays; i++) {
        tableval[i + week] = i + 1;
    }

    // 曜日
    str += '\t\t\t\t\t<tr class="week">';
    week = option_calendarstartweek;
    for (i = 0; i < 7; i++) {
        if (week === 0) {           str += '<th class="sunday">';
        } else if (week === 6) {        str += '<th class="saturday">';
        } else {                    str += '<th>';
        }
        str += list_week[week] + '</th>';
        week++;
        if (week >= 7) week = 0;
    }
    str += '</tr>\n';
    // 日付
    week = option_calendarstartweek;
    for (y = 0; y < calline; y++) {
        str += '\t\t\t\t\t<tr>';
        for (x = 0; x < 7; x++) {
            day = tableval[x + (y * 7)];
            // 指定日が祝日かどうか
            holiday = getHoliday(year, month, day);
            // 日記を書いた日である場合
            if (isDiary(year, month, day)) {
                str_date = year + formatNumber(month) + formatNumber(day);
                if (week === 0 || holiday > 0) {    str += '<td class="wrote-sunday">';
                } else if (week === 6) {            str += '<td class="wrote-saturday">';
                } else {                            str += '<td class="wrote">';
                }
                str += '<a href="#diary' + str_date + '">' + day + '</a></td>';
            } else {
                if (week === 0 || holiday > 0) {    str += '<td class="sunday">';
                } else if (week === 6) {            str += '<td class="saturday">';
                } else {                            str += '<td>';
                }
                str += day + '</td>';
            }
            week++;
            if (week >= 7) week = 0;
        }
        str += '</tr>\n';
    }
    str += '\t\t\t\t</table>\n';
    str += '\t\t\t</div>\n';
    str += '\t\t</section>\n';


    // ---------------------------------------------------------------------------------------------
    // サイドバーの日記ページ一覧のリンク
    // ---------------------------------------------------------------------------------------------
    str += '\t\t<section>\n';
    str += '\t\t\t<h4 class="caption-side">Archive</h4>\n';
    str += '\t\t\t<div class="archive">\n';
    for (i = 0; i < g_list_filename.length; i++) {
        // 日記ファイル名から年月を取り出して数値化
        year_tmp = Number(g_list_filename[i].substr(0, 4));
        month_tmp = Number(g_list_filename[i].substr(4, 2));
        // 指定月の日記件数を取得
        num = countMonthDiary(year_tmp, month_tmp);
        // 年が異なる場合
        if (year_old !== year_tmp) {
            year_old = year_tmp;
            if (i > 0) {
                str += '\t\t\t\t</ul>\n';
            }
            str += '\t\t\t\t<ul>\n';
            if (option_exportyeardiary) {
                str += '\t\t\t\t\t<li><a href="' + year_tmp + '.html">' + year_tmp + '年</a></li>\n';
            }
        }
        // カレントページ
        if (year_tmp === year && month_tmp === month) {
            str += '\t\t\t\t\t<li class="current">' + year_tmp + '年' + formatNumber(month_tmp) + '月';
        } else {
            // 「index.html」を最新日記の月にする場合、かつ、最新日記の年月の場合
            if (option_indexmode === 1 && (year_tmp === g_newdiaryyear && month_tmp === g_newdiarymonth)) {
                str += '\t\t\t\t\t<li><a href="index.html">' + year_tmp + '年' + formatNumber(month_tmp) + '月</a>';
            } else {
                str += '\t\t\t\t\t<li><a href="' + year_tmp + formatNumber(month_tmp) + '.html">' + year_tmp + '年' + formatNumber(month_tmp) + '月</a>';
            }
        }
        str += '<span class="count">（' + num + '件）</span></li>\n';
    }
    str += '\t\t\t\t</ul>\n';
    str += '\t\t\t</div>\n';
    str += '\t\t</section>\n';
    str += '\t</div>\n';

    // サイドバーのHTMLを返す
    return str;
}



// *************************************************************************************************
// 指定月の日記件数を取得
// *************************************************************************************************
function countMonthDiary(year, month)
{
    var cnt = 0;                                        // 日記の数
    var key = '';                                       // 連想配列のキー

    for (key in g_hash_diarytext) {
        if (g_hash_diarytext.hasOwnProperty(key)) {
            if (Number(key.substr(0, 4)) === year && Number(key.substr(4, 2)) === month) {
                cnt++;
            }
        }
    }
    return cnt;
}



// *************************************************************************************************
// ナビゲーションリンクHTML作成
// *************************************************************************************************
function createNavHTML(current_year, current_month)
{
    var str = '';                                       // HTML
    var filename = '';                                  // リンクファイル名
    var key = '';                                       // 連想配列のキー
    var i;                                              // ループ用


    str += '\t<nav>\n';

    // 連想配列を取り出す（連想配列は'2014':' 1  2  3  4  5  6  7  8  9  10  11  12 'のように区切られている）
    for (key in g_hash_nav) {
        if (g_hash_nav.hasOwnProperty(key)) {
            str += '\t\t<div class="nav-yearly">\n';
            if (option_exportyeardiary) {
                // カレントページの場合
                if (Number(key) === Number(current_year) && Number(current_month) === 0) {
                    str += '\t\t\t<span class="title current">' + key + '</span>\n';
                } else {
                    str += '\t\t\t<a class="title" href="' + key + '.html">' + key + '</a>\n';
                }
            } else {
                str += '\t\t\t<span class="title">' + key + '</span>\n';
            }
            // 12月分の文字列を作成
            for (i = 1; i <= 12; ++i) {
                // 月別のリンクファイル名を作成（「201401.html」など）
                filename = key + formatNumber(i) + '.html';
                // カレントページの場合
                if (Number(key) === current_year && i === current_month) {
                    str += '\t\t\t<span class="current">' + i + '</span>\n';
                // その月が文字列に含まれる場合
                } else if (g_hash_nav[key].match(' ' + i + ' ')) {
                    // 「index.html」を最新日記の月にする場合、かつ、最新日記の年月の場合
                    if (option_indexmode === 1 && (Number(key) === g_newdiaryyear && i === g_newdiarymonth)) {
                        str += '\t\t\t<a href="index.html">' + i + '</a>\n';
                    } else {
                        str += '\t\t\t<a href="' + filename + '">' + i + '</a>\n';
                    }
                } else {
                    str += '\t\t\t<span class="disabled">' + i + '</span>\n';
                }
            }
            str += '\t\t</div>\n';
        }
    }
    str += '\t</nav>\n';

    // ナビゲーションリンクのHTMLを返す
    return str;
}



// *************************************************************************************************
// 日記本文のHTML変換
// *************************************************************************************************
function convertHTML(str)
{
    // 行頭のタブを除去
    str = str.replace(/^\t/g, '');
    str = str.replace(/\n\t/g, '\n');
    // テキスト末尾の改行・空行を削除
    str = str.replace(/\n+$/g, '');

    // 日記本文のHTMLタグを有効にする場合
    if (option_enablehtmltag) {
        // タブを除去
        str = str.replace(/\t/g, '');
    // 日記本文のHTMLタグを有効にしない場合
    } else {
        // 文字参照の変換
        str = str.replace(/&/g, '&amp;');
        str = str.replace(/</g, '&lt;');
        str = str.replace(/>/g, '&gt;');
        // 連続するタブを半角スペースに変換
        str = str.replace(/\t+/g, '&nbsp;');
    }

    // 一行目を見出しにする場合、かつ、1文字目が「*」ではない場合
    if (option_diarytitle && str.substr(0, 1) !== '*') {
        str = '*' + str;
    }
    // 先頭に改行を付け足しておく（※あとの正規表現による置換をしやすくするため）
    str = '\n' + str;

    // 見出し
    str = str.replace(/\n\*(.*)/g, '\n<h3>$1</h3>');
    // 空の見出し（「<h3></h3>」や「<h3>半角スペース</h3>」）となった部分を削除
    str = str.replace(/\n<h3>\s*<\/h3>/g, '');

    // 動画
    str = str.replace(/\n\[\[(.[^:].*\.(3g2|3gp|3gp2|3gpp|asf|asx|avi|dat|dcr|div|divx|f4p|f4v|flc|fli|flv|m1v|m2t|m2ts|m2v|m3u|m4|m4a|m4b|m4p|m4r|m4v|mkv|mod|mov|mp2v|mp4|mp4v|mpe|mpeg|mpg|mts|ogm|qt|ram|rm|rmvb|swf|ts|tts|vdo|vg2|vgm|viv|vob|wax|wm|wmv|wpl|wrl|wvx))\]\]/gi,
        '\n<video controls src="data/$1" width="640"></video>');
    str = str.replace(/\n\[\[(.*\.(3g2|3gp|3gp2|3gpp|asf|asx|avi|dat|dcr|div|divx|f4p|f4v|flc|fli|flv|m1v|m2t|m2ts|m2v|m3u|m4|m4a|m4b|m4p|m4r|m4v|mkv|mod|mov|mp2v|mp4|mp4v|mpe|mpeg|mpg|mts|ogm|qt|ram|rm|rmvb|swf|ts|tts|vdo|vg2|vgm|viv|vob|wax|wm|wmv|wpl|wrl|wvx))\]\]/gi,
        '\n<video controls src="$1" width="640"></video>');
    // 音声
    str = str.replace(/\n\[\[(.[^:].*\.(aac|aif|aifc|aiff|au|cda|f4a|f4b|mid|midi|mp2|mp3|mpa|mpv2|oga|ogg|rmi|snd|wav|webm|wma))\]\]/gi,
        '\n<audio controls src="data/$1"></audio>');
    str = str.replace(/\n\[\[(.*\.(aac|aif|aifc|aiff|au|cda|f4a|f4b|mid|midi|mp2|mp3|mpa|mpv2|oga|ogg|rmi|snd|wav|webm|wma))\]\]/gi,
        '\n<audio controls src="$1"></audio>');
    // 画像（※その他の拡張子はすべて画像と見なす）
    str = str.replace(/\n\[\[(.[^:].*)\]\]/g, '\n<img src="data/$1" alt="">');
    str = str.replace(/\n\[\[(.*)\]\]/g, '\n<img src="$1" alt="">');

    // メールアドレス
    str = str.replace(/([\w\.\-]+@[\w\.\-]+)/gi, '<a class="link-mail" href="mailto:$1">$1</a>');
    // URL
    str = str.replace(/(https?:\/\/[\x21-\x7e]+)/gi, '<a class="link-url" href="$1">$1</a>');

    // ファイルパス
    str = str.replace(/\b([a-z]:\\[^\t\/:\*\?"<>\|\n&]*)/gi, '<a class="link-path" href="file:///$1">$1</a>');
    // <video><audio><img>タグの中に入り込んだ<a>タグ（フルパス）を元に戻しておく
    str = str.replace(/src="<a .*?">(.*?)<\/a>/gi, 'src="file:///$1');

    // テキスト先頭の改行を削除
    str = str.replace(/^\n+/g, '');

    // 改行を<br>に変換（<br>が必要ない行の改行を削除→改行を改行タグに変換→<br>が必要ない行の改行を戻す）
    str = str.replace(/<\/(h[1-6]|div|p|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)>\n/gi, '</$1>');
    str = str.replace(/<(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)>\n/gi, '<$1>');
    str = str.replace(/\n/g, '<br>\n');
    str = str.replace(/<\/(h[1-6]|div|p|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)>/gi, '</$1>\n');
    str = str.replace(/<(div|p|hr|blockquote|ul|ol|dl|table|tr)>/gi, '<$1>\n');

    // <p>タグ変換（連続した改行を<br><br><br>...とせずに<p>タグで囲んできちんと変換）
    if (option_ptagconvert) {
        // ブロックレベル要素の前後に<p>を挿入
        str = str.replace(/<\/(h[1-6]|div|p|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)>\n/gi, '</$1>\n<p>\n');
        str = str.replace(/<br>\n<(h[1-6]|div|p|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)/gi, '\n</p>\n<$1');
        // 2個以上連続する<br>を<p>に変換
        str = str.replace(/(<br>\n){2,}/gi, '\n</p>\n<p>\n');

        // <br>だけの行を削除
        str = str.replace(/\n<br>/gi, '');
        // 先頭と末尾に<p>を追加
        str = '<p>\n' + str + '\n</p>';
        // 空行を削除
        str = str.replace(/\n+/gi, '\n');

        // テキストがない部分（<p>\n</p>の部分）を削除
        str = str.replace(/<p>\n<\/p>\n/gi, '');

        // ブロックレベル要素の間の<p>と</p>を削除
        str = str.replace(/<(\/?)(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)(.*?)>\n<\/?p>\n<(\/?)(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)\b/gi, '<$1$2$3>\n<$4$5');
        str = str.replace(/<(\/?)(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)(.*?)>\n<\/?p>\n<(\/?)(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)\b/gi, '<$1$2$3>\n<$4$5');
        // ブロックレベル要素の前後の<p>と</p>を削除
        str = str.replace(/^<p>\n<(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)/gi, '<$1');
        str = str.replace(/<\/(h[1-6]|div|p|hr|blockquote|ul|ol|li|dl|dt|dd|table|tr|th|td)>\n<\/p>/gi, '</$1>\n');

        // インデントする（<p>～</p>内をブロックインデント）
        if (true) {
            // タグで始まらない行をタブインデント
            str = str.replace(/\n([^<])/gi, '\n\t$1');
            // <p>の中のインライン要素の行をタブインデント
            str = str.replace(/\n<(img|audio|video|a|b|i|u|s|em|strong|ins|del|mark|sup|sub|small|li|dt|dd|tr|th|td)\b(.*?)>/gi, '\n\t<$1$2>');
            str = str.replace(/\n<\/(dt|dd|tr|th|td)>/gi, '\n\t</$1>');
        // インデントしない（<p>～</p>内を一行にまとめる）
        } else {
            str = str.replace(/(<p>)\n/gi, '$1');
            str = str.replace(/\n(<\/p>)/gi, '$1');
            str = str.replace(/(<br>)\n/gi, '$1');
        }
    }

    // テキスト末尾の改行・空行を削除
    str = str.replace(/\n+$/g, '');

    return str;
}



// *************************************************************************************************
// 目次ページ（index.html）の出力
// 【引数】
//      path                日記フォルダのパス（例「c:\wDiary\MyDiary」など※最後の\は付けない）
//      foldername          日記フォルダの名前のパス（日記の名前）
// *************************************************************************************************
function createIndexPage(path, foldername)
{
    var str = '';                                       // HTML全体の文字列
    var html_navigation = '';                           // 月別のリンク
    var year, month, day;                               // 年月日
    var year_old = 0, month_old = 0;                    // 直前の年と月
    var str_line = '';                                  // 一行の文字列
    var i = 0;                                          // ループ用
    var key = '';                                       // 連想配列のキー


    // ナビゲーションリンクHTML作成
    html_navigation = createNavHTML(0, 0);

    str += '<!DOCTYPE html>\n';
    str += '<html>\n';
    str += '<head>\n';
    str += '\t<meta charset="' + option_charcode + '">\n';
    str += '\t<meta name="viewport" content="initial-scale=1.0">\n';
    str += '\t<title>' + foldername + '</title>\n';
    str += '\t<link rel="stylesheet" href="css/style.css">\n';
    str += '</head>\n';
    str += '<body>\n';
    str += '\t<header>\n';
    str += '\t\t<h1 class="diary-name">' + ((option_indexmode >= 0) ? ('<a href="index.html">' + foldername + '</a>') : foldername) + '</h1>\n';
    str += '\t</header>\n';
    str += '\n';
    str += (option_navigationtop ? (html_navigation + '\n') : '');      // ナビゲーションリンク
    str += '\t<div class="index">\n';

    for (key in g_hash_diarytext) {
        if (g_hash_diarytext.hasOwnProperty(key)) {
            // キーから年月日を取り出して数値化
            year = Number(key.substr(0, 4));
            month = Number(key.substr(4, 2));
            day = Number(key.substr(6, 2));
            if (month !== month_old) {
                if (i > 0) str += '\t\t\t</div>\n';
            }
            if (year !== year_old) {
                if (i > 0) str += '\t\t</div>\n';
                if (option_exportyeardiary) {
                    str += '\t\t<h2 class="caption-year"><a href="' + year + '.html">' + year + '年</a></h2>\n';
                } else {
                    str += '\t\t<h2 class="caption-year">' + year + '年</h2>\n';
                }
                str += '\t\t<div class="indent">\n';
            }
            if (month !== month_old) {
                str += '\t\t\t<h3 class="caption-month"><a href="' + year + formatNumber(month) + '.html">' + month + '月</a></h3>\n';
                str += '\t\t\t<div class="indent">\n';
            }
            // 文字列から一行目を取り出す
            str_line = getTopLine(g_hash_diarytext[key]);
            // HTML作成
            str += '\t\t\t\t<div class="topic">' + year + '-' + formatNumber(month) + '-' + formatNumber(day) + ' <a href="' + year + formatNumber(month) + '.html#diary';
            str += year + formatNumber(month) + formatNumber(day) + '">' + str_line + '</a></div>\n';

            month_old = month;
            year_old = year;
            i++;
        }
    }
    if (i > 0) {
        str += '\t\t\t</div>\n';
        str += '\t\t</div>\n';
    }

    str += '\t</div>\n';
    str += '\n';
    str += (option_navigationbottom ? (html_navigation + '\n') : '');       // ナビゲーションリンク
    str += '\t<footer>\n';
    str += '\t\t<div class="copyright">\n';
    str += '\t\t\t&copy; ' + g_now.getFullYear() + ((option_indexmode >= 0) ? (' <a href="index.html">' + foldername + '</a>') : (' ' + foldername)) + ' All rights reserved.<br>\n';
    str += '\t\t\tPowered by &copy; <a href="http://www.cc9.ne.jp/~pappara/wdiary.html">wDiary</a>\n';
    str += '\t\t</div>\n';
    str += '\t</footer>\n';
    str += '\n';
    str += '</body>\n';
    str += '</html>\n';

    // HTML最終整形
    str = formatHTML(str);

    // ファイル書き込み
    return saveFile(path + '\\index.html', str);
}



// *************************************************************************************************
// HTML最終整形
// *************************************************************************************************
function formatHTML(str)
{
    var i;                                              // スペースインデント用
    var tmp = '';                                       // スペースインデント用

    // HTMLのインデントなし
    if (option_indent < 0) {
        // タブを除去
        str = str.replace(/\t/g, '');
        // 空行の削除
        str = str.replace(/\n+/g, '\n');
    // HTMLのタブインデント
    } else if (option_indent === 0) {
        // そのまま
    // HTMLのスペースインデント
    } else {
        for (i = 0; i < option_indent; ++i) tmp += ' ';
        str = str.replace(/\t/g, tmp);
    }

    // 閉じスラッシュ
    if (option_closetagslash) {
        str = str.replace(/<(br|hr)>/gi, '<$1 />');
        str = str.replace(/<(meta|link|img) (.*?)>/gi, '<$1 $2 />');
    }

    // 改行コード変換
    if (option_linecode !== '\n') {
        str = str.replace(/\n/g, option_linecode);
    }
    return str;
}



// *************************************************************************************************
// 文字列から一行目を取り出す
// *************************************************************************************************
function getTopLine(str)
{
    // 行頭のタブを除去
    str = str.replace(/^\t+/g, '');
    // 行頭の空行を削除
    str = str.replace(/^\n+/g, '');
    // 文字参照の変換
    str = str.replace(/&/g, '&amp;');
    str = str.replace(/</g, '&lt;');
    str = str.replace(/>/g, '&gt;');
    // 一行目を取り出す
    str = str.split('\n')[0];
    // 行頭の「*」を削除
    str = str.replace(/^\*/g, '');
    return str;
}



// *************************************************************************************************
// ファイル読み込み
// *************************************************************************************************
function loadFile(path)
{
    var str = '';                                       // 読み込んだテキスト

    // （ADODB.Stream版）
    g_as.type = 2;                      // データのタイプ（1:バイナリ 2:テキスト）
    g_as.charset = 'unicode';           // 文字セット（'utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp'など）
    g_as.open();
    g_as.loadFromFile(path);
    str = g_as.readText();
    g_as.close();
    return str;

    // （FileSystemObject版）
//  var hFile = g_fs.OpenTextFile(path, 1, true, -1);
//  str = hFile.ReadAll();
//  hFile.Close();
//  return str;
}



// *************************************************************************************************
// ファイル書き込み
// *************************************************************************************************
function saveFile(path, str)
{
    // （ADODB.Stream版）
    g_as.Type = 2;                      // データのタイプ（1:バイナリ 2:テキスト）
    g_as.charset = option_charcode;     // 文字セット（'utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp'など）
    g_as.Open();
    g_as.WriteText(str);
    try {
        g_as.SaveToFile(path, 2);       // 第二引数（1:ファイルが存在すればエラー 2:上書き）
    } catch (e) {
        g_as.Close();
        return false;
    }
    g_as.Close();
    return true;

    // （FileSystemObject版）
//  var hFile = g_fs.OpenTextFile(path, 2, true, -1);
//  hFile.Write(str);
//  hFile.Close();
}



// *************************************************************************************************
// 数値を2桁の文字列に揃える
// *************************************************************************************************
function formatNumber(num)
{
    return String((num < 10) ? ('0' + num) : num);
}



// *************************************************************************************************
// 連想配列をキーでソート
// *************************************************************************************************
function hashSort(obj, flag_reverse)
{
    var ret = {};                                       // 戻り値用の連想配列
    var arr = [];                                       // キーを格納する配列
    var key = '';                                       // 連想配列のキー
    var i;                                              // ループ用

    // キーを配列に取り出してソート
    for (key in obj) if (obj.hasOwnProperty(key)) arr.push(key);
    arr.sort();
    if (flag_reverse) arr.reverse();
    // 新しい連想配列に入れ直す
    for (i = 0; i < arr.length; i++) ret[arr[i]] = obj[arr[i]];
    return ret;
}



// *************************************************************************************************
// 連想配列の結合
// *************************************************************************************************
function hashMerge(hash1, hash2)
{
    var ret = {};                                       // 戻り値用の連想配列
    var key = '';                                       // 連想配列のキー

    for (key in hash1) if (hash1.hasOwnProperty(key)) ret[key] = hash1[key];
    for (key in hash2) if (hash2.hasOwnProperty(key)) ret[key] = hash2[key];
    return ret;
}



// *************************************************************************************************
// 指定日が祝日（休日）かどうか
// 【戻り値】
//      0: 平日
//      1: 元日(1月1日)
//      2: 成人の日(1月第2月曜日)
//      3: 建国記念の日(2月11日)
//      4: 春分の日(春分日(3月20日)※)
//      5: 昭和の日(4月29日)
//      6: 憲法記念日(5月3日)
//      7: みどりの日(5月4日)
//      8: こどもの日(5月5日)
//      9: 海の日(7月第3月曜日)
//      10: 山の日(8月11日)
//      11: 敬老の日(9月第3月曜日)
//      12: 秋分の日(秋分日(9月23日)※)
//      13: 体育の日(10月第2月曜日※2020年以降体育の日からスポーツの日へ)
//      14: 文化の日(11月3日)
//      15: 勤労感謝の日(11月23日)
//      16: 天皇誕生日(2019年まで12月23日)(2020年から2月23日)
//      17: 振替休日
//      18: 国民の休日
//      19: 新天皇即位（2019/05/01）
//      20: 即位礼正殿の儀（2019/10/22）
//      21: スポーツの日
// *************************************************************************************************
function getHoliday(year, month, day)
{
    var week;                                           // 曜日（0:日 1:月 2:火 3:水 4:木 5:金 6:土）
    var specialday;                                     // 特別な日（春分の日、秋分の日）

    if (month === 1) {
        if (day === 1) return 1;                                                // 1: 元日
        if (day === 2 && (0 === getWeekDay(year, month, 1))) return 17;         // 17: 振替休日
        if (Math.floor((day - 1) / 7) === 1) {
            if (1 === getWeekDay(year, month, day)) return 2;                   // 2: 成人の日
        }
    } else if (month === 2) {
        if (day === 11) return 3;                                               // 3: 建国記念の日
        if (day === 12 && (0 === getWeekDay(year, month, 11))) return 17;       // 17: 振替休日
        if (year >= 2020) {
            if (day === 23) return 16;                                          // 16: 天皇誕生日
            if (day === 24 && (0 === getWeekDay(year, month, 23))) return 17;   // 17: 振替休日
        }
    } else if (month === 3) {
        if (day >= 19 && day <= 22) {
            // 春分の日を取得
            specialday = getVernalEquinoxDay(year);
            if (day === specialday) return 4;                                   // 4: 春分の日（３月２１日ごろ）
            if (day === specialday + 1) {
                if (getWeekDay(year, month, specialday) === 0) return 17;       // 17: 振替休日
            }
        }
    } else if (month === 4) {
        if (day === 29) return 5;                                               // 5: 昭和の日
        if (day === 30 && (0 === getWeekDay(year, month, 29))) return 17;       // 17: 振替休日
        if (year === 2019 && day === 30) return 18;                             // 18: 国民の休日
    } else if (month === 5) {
        if (day === 3) return 6;                                                // 6: 憲法記念日
        if (day === 4) return 7;                                                // 7: みどりの日
        if (day === 5) return 8;                                                // 8: こどもの日
        if (day === 6) {
            week = getWeekDay(year, month, 3);
            if (week === 5 || week === 6 || week === 0) return 17;              // 17: 振替休日
        }
        if (year === 2019) {
            if (day === 1) return 19;                                           // 19: 新天皇即位
            if (day === 2) return 18;                                           // 18: 国民の休日
        }
    } else if (month === 7) {
        if (year === 2020) {
            if (day === 23) return 9;                                           // 9: 海の日
            if (day === 24) return 21;                                          // 21: スポーツの日
        } else if (year === 2021) {
            if (day === 22) return 9;                                           // 9: 海の日
            if (day === 23) return 21;                                          // 21: スポーツの日
        } else {
            if (Math.floor((day - 1) / 7) === 2) {
                if (1 === getWeekDay(year, month, day)) return 9;               // 9: 海の日
            }
        }
    } else if (month === 8 && year >= 2016) {
        if (year === 2020) {
            if (day === 10) return 10;                                          // 10: 山の日
        } else if (year === 2021) {
            if (day === 8) return 10;                                           // 10: 山の日
            if (day === 9) return 17;                                           // 17: 振替休日
        } else {
            if (day === 11) return 10;                                          // 10: 山の日
            if (day === 12 && (0 === getWeekDay(year, month, 11))) return 17;   // 17: 振替休日
        }
    } else if (month === 9) {
        if (Math.floor((day - 1) / 7) === 2) {
            if (getWeekDay(year, month, day) === 1) return 11;                  // 11: 敬老の日
        }
        if (day >= 21 && day <= 25) {
            // 秋分の日を取得
            specialday = getAutumnalEquinoxDay(year);
            if (day === specialday) return 12;                                  // 12: 秋分の日
            week = getWeekDay(year, month, specialday);
            if (day === specialday + 1) {
                if (week === 0) return 17;                                      // 17: 振替休日
            }
            if (week === 3 && day === specialday - 1) return 18;                // 18: 国民の休日
        }
    } else if (month === 10) {
        if (year === 2021) return 0;                                            // なし
        if (year != 2020 && (Math.floor((day - 1) / 7) === 1)) {
            if (1 === getWeekDay(year, month, day)) {
                return (year < 2020) ? 13 : 21;                                 // 13: 体育の日、21: スポーツの日
            }
        }
        if (year === 2019 && day === 22) return 20;                             // 20: 即位礼正殿の儀
    } else if (month === 11) {
        if (day === 3) return 14;                                               // 14: 文化の日
        if (day === 4) {
            if (getWeekDay(year, month, 3) === 0) return 17;                    // 17: 振替休日
        }
        if (day === 23) return 15;                                              // 15: 勤労感謝の日（１１月２３日）
        if (day === 24) {
            if (getWeekDay(year, month, 23) === 0) return 17;                   // 17: 振替休日
        }
    } else if (month === 12) {
        if (year <= 2018) {
            if (day === 23) return 16;                                          // 16: 天皇誕生日
            if (day === 24) {
                if (getWeekDay(year, month, 23) === 0) return 17;               // 17: 振替休日
            }
        }
    }
    return 0;
}



// *************************************************************************************************
// 日付の曜日を求める（0:日 1:月 2:火 3:水 4:木 5:金 6:土）
// *************************************************************************************************
function getWeekDay(year, month, day)
{
    return new Date(year, month - 1, day).getDay();
}



// *************************************************************************************************
// 春分の日を取得（1900年～2150年まで対応）
// *************************************************************************************************
function getVernalEquinoxDay(year)
{
    if (year <= 1979) {         return Math.floor(20.8357 + (0.242194 * (year - 1980)) - Math.floor((year - 1983) / 4));
    } else if (year <= 2099) {  return Math.floor(20.8431 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    } else if (year <= 2150) {  return Math.floor(21.8510 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    return 0;
}



// *************************************************************************************************
// 秋分の日を取得（1900年～2150年まで対応）
// *************************************************************************************************
function getAutumnalEquinoxDay(year)
{
    if (year <= 1979) {         return Math.floor(23.2588 + (0.242194 * (year - 1980)) - Math.floor((year - 1983) / 4));
    } else if (year <= 2099) {  return Math.floor(23.2488 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    } else if (year <= 2150) {  return Math.floor(24.2488 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    return 0;
}
