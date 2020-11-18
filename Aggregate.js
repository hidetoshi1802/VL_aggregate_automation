// 宿題FB數の自動集計・入力スクリプト
function aggregate() {
    // 講師代カウントシート
    var this_ss = SpreadsheetApp.getActiveSpreadsheet();
    // 集計対象月
    var target_month = this_ss.getRange('A1').getValue();
    var disp_target_date = Utilities.formatDate(target_month, 'JST', 'yyyy年MM月');

    var result = Browser.msgBox('処理対象月：' + disp_target_date, '集計結果を更新してよろしいですか？\\n※既存の数値は上書きされます。', Browser.Buttons.OK_CANCEL);
    if (result == "cancel") {
        Logger.log("canceled...");
        return;
    }

    // 講師名
    var teacher_name = this_ss.getActiveSheet().getName();
    // 集計結果入力列
    var target_input_col = findCol(this_ss, target_month);

    var vlo_periods_first_row = 4;

    if (!target_input_col) {
        Browser.msgBox('Error', '入力対象の月が存在しません。', Browser.Buttons.OK);
        return;
    }

    var last_row = this_ss.getActiveSheet().getLastRow();
    // 集計対象VLO期
    var vlo_periods = getVLOPeriods(this_ss, vlo_periods_first_row, last_row);

    if (!vlo_periods) {
        Browser.msgBox('Error', '集計対象のVLO期が存在しません。', Browser.Buttons.OK);
        return;
    }

    // 集計処理
    var result = getResult(target_month, teacher_name, vlo_periods);

    var cnt_result = Object.keys(result).length;

    // 入力処理
    var result_key = Object.keys(result);
    var vlo_period_col = 'A';
    var values_col_A = this_ss.getActiveSheet().getRange(vlo_period_col + '1' + ':' + vlo_period_col + last_row).getValues();

    for (var i = 0; i < result_key.length; i++) {
        var vlo_period_name = result_key[i];
        var start_row = null;

        for (var i1 = 0; i1 < values_col_A.length; i1++) {

            if (values_col_A[i1][0] == vlo_period_name) {
                start_row = i1 + 1;
                break;
            }

        }

        var result_values = result[vlo_period_name];

        for (var i1 = 0; i1 < Object.keys(result_values).length; i1++) {
            var input_val = null;
            var val = result_values[i1 + 1];


            if (val == 'Undelivered') {
                continue;
            } else {

                if (!val) {
                    input_val = 0;
                } else {
                    input_val = val.length;
                }

                var input_cell = this_ss.getActiveSheet().getRange(start_row + i1, target_input_col);
                input_cell.setValue(input_val);
            }

        }

    }

    Browser.msgBox('処理対象月：' + disp_target_date, '集計結果の更新が完了しました。', Browser.Buttons.OK);
    return;

}


//　集計結果入力列を取得
function findCol(this_ss, target_month) {
    // 年月の行
    var row_num_month = 4;

    var last_col = this_ss.getLastColumn();
    var month = Utilities.formatDate(target_month, "JST", "YYYY/MM/dd")
    var values = this_ss.getActiveSheet().getRange(row_num_month, 1, 1, last_col).getValues();
    var new_values = [];

    for (var i = 0; i < values.length; i++) {
        var newValues = values[i].map(
            function (x) {
                var type = Object.prototype.toString.call(x);
                if (type == "[object Date]") {
                    return x = Utilities.formatDate(x, 'JST', 'yyyy/MM/dd');
                }
            });
        new_values[i] = newValues;
    }

    var str_target_month = Utilities.formatDate(target_month, 'JST', 'yyyy/MM/dd');
    var col_num = new_values[0].indexOf(str_target_month);

    if (col_num >= 0) {
        return col_num + 1;
    } else {
        return false;
    }

}


// 集計対象のVLO期を取得
function getVLOPeriods(this_ss, vlo_periods_first_row, last_row) {
    var values = this_ss.getActiveSheet().getRange(vlo_periods_first_row, 1, last_row, 1).getValues();
    var vlo_periods = [];

    values.map(function (v1) {

        if (v1) {
            v1.map(function (v2) {

                if (v2.match('VLO.*?期')) {
                    vlo_periods.push(v2);
                }

            });
        }

    });

    return vlo_periods;
}


// 「VLO参加者・宿題フィードバック担当講師」シートから集計
function getResult(target_month, teacher_name, vlo_periods) {
    // VLO参加者・宿題フィードバック担当講師シート
    var target_ss = SpreadsheetApp.openById('15zfjF0pRIFs7MeEDwEiwNtLSEYlp05pmCzlvEEb5PAM');
    var result = {};
    var str_target_date = Utilities.formatDate(target_month, 'JST', 'yyyy/MM/dd');

    // 集計処理
    for (var i0 = 0; i0 < vlo_periods.length; i0++) {
        var sheet_name = vlo_periods[i0];
        var target_sheet = target_ss.getSheetByName(sheet_name);
        var last_row = target_sheet.getLastRow();
        var last_column = target_sheet.getLastColumn();
        // 講師名の列{●講:列}
        var column_name = null;

        // VLO4期の途中まで清算方法が違うため、集計対象となる講を指定
        // 講師名の列もずれているため、指定
        if (sheet_name == 'VLO1期') {
            // VLO1期は9講以降
            column_name = { 9: 'AG', 10: 'AM', 11: 'AS', 12: 'AY' };
        } else if (sheet_name == 'VLO2期') {
            // VLO2期は7講以降
            column_name = { 7: 'AE', 8: 'AK', 9: 'AQ', 10: 'AW', 11: 'BC', 12: 'BI' };
        } else if (sheet_name == 'VLO3期') {
            // VLO3期は5講以降
            column_name = { 5: 'Y', 6: 'AE', 7: 'AK', 8: 'AQ', 9: 'AW', 10: 'BC', 11: 'BI', 12: 'BO' };
        } else if (sheet_name == 'VLO4期') {
            // VLO4期は3講以降
            column_name = { 3: 'O', 4: 'U', 5: 'AA', 6: 'AG', 7: 'AM', 8: 'AS', 9: 'AY', 10: 'BE', 11: 'BK', 12: 'BQ' };
        } else if (sheet_name == 'VLO8期') {
            // VLO8期は列がずれているため、個別指定
            column_name = { 1: 'B', 2: 'H', 3: 'N', 4: 'T', 5: 'Z', 6: 'AF', 7: 'AL', 8: 'AR', 9: 'BD', 10: 'BJ', 11: 'BP', 12: 'BV' };
        } else {
            // 上記以外
            column_name = { 1: 'B', 2: 'H', 3: 'N', 4: 'T', 5: 'Z', 6: 'AF', 7: 'AL', 8: 'AR', 9: 'AX', 10: 'BD', 11: 'BJ', 12: 'BP' };
        }

        var keys = Object.keys(column_name);

        var period_result = {};

        for (var i1 = 0; i1 < keys.length; i1++) {
            var column = column_name[keys[i1]];
            var values = target_sheet.getRange(column + '1' + ':' + column + last_row).getValues();
            var delivery_date = target_sheet.getRange(column + '4').getValue();
            var array = [];
            var row_num = [];

            for (var i2 = 0; i2 < values.length; i2++) {
                array.push(values[i2][0]);
            }

            for (var i3 = 0; i3 < array.length; i3++) {
                var val = array[i3];
                var row = i3 + 1;

                if (array[i3] == teacher_name) {
                    // FB実施日
                    var cell_date = target_sheet.getRange(column + row).offset(0, 5);
                    var execution_date = cell_date.getValue();
                    if (Object.prototype.toString.call(execution_date) == "[object Date]") {
                        var str_date = Utilities.formatDate(execution_date, 'JST', 'yyyy/MM/dd')
                        if (Moment.moment(str_target_date).isSame(str_date, 'year') && Moment.moment(str_target_date).isSame(str_date, 'month')) {
                            row_num.push(row);
                        }
                    }
                }

            }

            if (row_num.length > 0) {
                period_result[i1 + 1] = row_num;
            } else {

                if (delivery_date && Object.prototype.toString.call(delivery_date) == "[object Date]") {
                    var str_delivery_date = Utilities.formatDate(delivery_date, 'JST', 'yyyy/MM/dd');
                    var today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

                    if (Moment.moment(str_delivery_date).isBefore(today)) {
                        period_result[i1 + 1] = [];
                    } else {
                        period_result[i1 + 1] = 'Undelivered';
                    }

                }
            }

            if (!delivery_date) {
                period_result[i1 + 1] = 'Undelivered';
            }

        }

        if (Object.keys(period_result).length) {
            result[sheet_name] = period_result;
        }

    }

    return result;
}
