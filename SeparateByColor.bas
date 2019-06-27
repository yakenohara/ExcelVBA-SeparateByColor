Attribute VB_Name = "SeparateByColor"
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
' 選択範囲を色分けする
'
Sub SeparateByColor()
    
    '<Color Setting>-------------------------------------------------
    '
    '|Interior.Color RGB Value | Font.Color RGB Value    |
    '|   Array(rrr,ggg,bbb)    |   Array(rrr,ggg,bbb)    |
    '
    arr_colors = Array( _
        Array(Array(28, 156, 116), Array(255, 255, 255)), _
        Array(Array(220, 92, 4), Array(255, 255, 255)), _
        Array(Array(116, 116, 180), Array(255, 255, 255)), _
        Array(Array(228, 44, 140), Array(255, 255, 255)), _
        Array(Array(100, 164, 28), Array(255, 255, 255)), _
        Array(Array(228, 172, 4), Array(255, 255, 255)), _
        Array(Array(163, 115, 28), Array(255, 255, 255)), _
        Array(Array(134, 122, 128), Array(255, 255, 255)) _
    )
    
    '------------------------------------------------</Color Setting>
    
    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して値の書き込みを行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
    
    '実行確認
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    'シート選択状態チェック
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '<スタイル削除>------------------------------------------------
    With Selection.Interior '背景カラーの解除
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font  'フォントカラーの解除
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Borders.LineStyle = xlNone  '枠線の解除
    '-----------------------------------------------</スタイル削除>
    
    '<Color setting の collection 化>------------------------------------------------
    
    Set collection_colors = New Collection
    
    For long_1d_counter = LBound(arr_colors, 1) To UBound(arr_colors, 1)
        
        arr_tmp = arr_colors(long_1d_counter)
        
        long_tmp_base_1 = LBound(arr_tmp, 1)
        
        arr_interior_color = arr_tmp(long_tmp_base_1)
        long_tmp_base_2 = LBound(arr_interior_color)
        
        long_interior_color = RGB( _
            arr_interior_color(long_tmp_base_2), _
            arr_interior_color(long_tmp_base_2 + 1), _
            arr_interior_color(long_tmp_base_2 + 2) _
        )
        
        arr_font_color = arr_tmp(long_tmp_base_1 + 1)
        long_tmp_base_2 = LBound(arr_font_color)
        
        long_font_color = RGB( _
            arr_font_color(long_tmp_base_2), _
            arr_font_color(long_tmp_base_2 + 1), _
            arr_font_color(long_tmp_base_2 + 2) _
        )
        
        Dim dict_color As Object
        Set dict_color = CreateObject("Scripting.Dictionary")
        
        With dict_color
            .Add "prop_long_interior_color", long_interior_color
            .Add "prop_long_font_color", long_font_color
        End With
        
        collection_colors.Add dict_color
        
    Next long_1d_counter
    
    '-----------------------------------------------</Color setting の collection 化>
    
    '<Color Dictionary の生成>------------------------------------------------
    
    Dim dict_colors As Object
    Set dict_colors = CreateObject("Scripting.Dictionary")
    
    Dim variant_2d_arr As Variant
    
    If Selection.CountLarge = 1 Then '対象セルが1だけの場合
        
        ReDim variant_2d_arr(1, 1) '1つだけの要素を持った2次元配列として定義
        variant_2d_arr(1, 1) = Selection.Value2
        
    Else
        
        '対象範囲がUsedRange内に収まるようにトリミングして2次元配列化
        Set range_tmp = trimWithUsedRange(Selection)
        variant_2d_arr = range_tmp.Value2
    
    End If
    
    long_lower_index_1d = LBound(variant_2d_arr, 1)
    long_upper_index_1d = UBound(variant_2d_arr, 1)
    long_lower_index_2d = LBound(variant_2d_arr, 2)
    long_upper_index_2d = UBound(variant_2d_arr, 2)
    
    long_base_index_1d = long_lower_index_1d - 0
    long_base_index_2d = long_lower_index_2d - 0
    
    long_last_index_no_of_color_collection = collection_colors.Count
    
    long_selecting_idx_no_of_color_collection = long_last_index_no_of_color_collection
    
    For long_1d_counter = long_lower_index_1d To long_upper_index_1d '行ループ
        For long_2d_counter = long_lower_index_2d To long_upper_index_2d '列ループ
        
            Set range_tmp = Cells( _
                Selection(1).Row + long_1d_counter - long_base_index_1d, _
                Selection(1).Column + long_2d_counter - long_base_index_2d _
            )
            
            '結合セルの場合は、結合セルの左上のみ捜査対象にする
            If (range_tmp.Address = range_tmp.MergeArea.Cells(1, 1).Address) Then
                
                
                '空のセルは無視する
                If (Not (IsEmpty(range_tmp.Value2))) Then
                
                    str_prop_name = "prop_" & TypeName(range_tmp.Value2) & "_" & CStr(range_tmp.Value2)
                    
                    'dictionary に追加していない場合に追加する
                    If Not (dict_colors.Exists(str_prop_name)) Then
                    
                        long_selecting_idx_no_of_color_collection = long_selecting_idx_no_of_color_collection + 1
                        If (long_selecting_idx_no_of_color_collection > long_last_index_no_of_color_collection) Then
                            long_selecting_idx_no_of_color_collection = 1
                        End If
                        
                        dict_colors.Add str_prop_name, long_selecting_idx_no_of_color_collection
                        
                    End If

                End If
            
            End If
            
        Next long_2d_counter
    Next long_1d_counter
    
    '-----------------------------------------------</Color Dictionary の生成>
    
    '<Coloring>------------------------------------------------
    
    For long_1d_counter = long_lower_index_1d To long_upper_index_1d '行ループ
        For long_2d_counter = long_lower_index_2d To long_upper_index_2d '列ループ
        
            Set range_tmp = Cells( _
                Selection(1).Row + long_1d_counter - long_base_index_1d, _
                Selection(1).Column + long_2d_counter - long_base_index_2d _
            )
            
            '結合セルの場合は、結合セルの左上のみ捜査対象にする
            If (range_tmp.Address = range_tmp.MergeArea.Cells(1, 1).Address) Then
                
                
                '空のセルは無視する
                If (Not (IsEmpty(range_tmp.Value2))) Then
                
                    str_prop_name = "prop_" & TypeName(range_tmp.Value2) & "_" & CStr(range_tmp.Value2)
                    
                    Set dict_color_specifier = collection_colors.Item(dict_colors.Item(str_prop_name))
                    
                    With range_tmp.Interior '背景カラー
                        .Color = dict_color_specifier.Item("prop_long_interior_color")
                    End With
                    With range_tmp.Font  'フォントカラー
                        .Color = dict_color_specifier.Item("prop_long_font_color")
                    End With
                    
                    'dictionary に追加していない場合に追加する
                    If Not (dict_colors.Exists(str_prop_name)) Then
                    
                        long_selecting_idx_no_of_color_collection = long_selecting_idx_no_of_color_collection + 1
                        If (long_selecting_idx_no_of_color_collection > long_last_index_no_of_color_collection) Then
                            long_selecting_idx_no_of_color_collection = 1
                        End If
                        
                        dict_colors.Add str_prop_name, long_selecting_idx_no_of_color_collection
                        
                    End If

                End If
            
            End If
            
        Next long_2d_counter
    Next long_1d_counter
    
    '-----------------------------------------------</Coloring>
    
    '<同色が隣接しているかチェック>------------------------------------------------
    
    For long_1d_counter = long_lower_index_1d To long_upper_index_1d '行ループ
        For long_2d_counter = long_lower_index_2d To long_upper_index_2d '列ループ
        
            Dim range_tmp_me As Range
            Set range_tmp_me = Cells( _
                Selection(1).Row + long_1d_counter - long_base_index_1d, _
                Selection(1).Column + long_2d_counter - long_base_index_2d _
            )
            
            variant_val2_of_me = range_tmp_me.MergeArea.Cells(1, 1).Value2 '結合セルの場合は左上セル値を使用
                
            '空のセルは無視する
            If (Not (IsEmpty(variant_val2_of_me))) Then
        
                str_me_prop_name = "prop_" & TypeName(variant_val2_of_me) & "_" & CStr(variant_val2_of_me)
                long_me_index_no_of_color_collection = dict_colors.Item(str_me_prop_name)
                
                Dim range_tmp_vs As Range
                
                '<vs 右セル>------------------------------------------------
                Set range_tmp_vs = Cells( _
                    Selection(1).Row + long_1d_counter - long_base_index_1d, _
                    Selection(1).Column + long_2d_counter - long_base_index_2d + 1 _
                )
                
                ret = checkSameColor( _
                    range_me:=range_tmp_me, _
                    range_vs:=range_tmp_vs, _
                    str_me_prop_name:=str_me_prop_name, _
                    long_border_specifier:=xlEdgeRight, _
                    dict_color_setting:=dict_colors, _
                    long_me_collor_no:=long_me_index_no_of_color_collection _
                )
                '-----------------------------------------------</vs 右セル>
                
                '<vs 下セル>------------------------------------------------
                Set range_tmp_vs = Cells( _
                    Selection(1).Row + long_1d_counter - long_base_index_1d + 1, _
                    Selection(1).Column + long_2d_counter - long_base_index_2d _
                )
                
                ret = checkSameColor( _
                    range_me:=range_tmp_me, _
                    range_vs:=range_tmp_vs, _
                    str_me_prop_name:=str_me_prop_name, _
                    long_border_specifier:=xlEdgeBottom, _
                    dict_color_setting:=dict_colors, _
                    long_me_collor_no:=long_me_index_no_of_color_collection _
                )
                '-----------------------------------------------</vs 下セル>
                
            End If
            
        Next long_2d_counter
    Next long_1d_counter
    
    '-----------------------------------------------</同色が隣接しているかチェック>
    
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub

'
' セル参照範囲が UsedRange 範囲に収まるようにトリミングする
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '指定範囲の右下位置の取得
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRangeの右下位置の取得
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    'トリミング
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '格納して終了
    Set trimWithUsedRange = ret
    
End Function

'
' 隣接セルの着色が同じ場合に枠線を引く
'
Private Function checkSameColor( _
    ByRef range_me As Range, _
    ByRef range_vs As Range, _
    ByVal str_me_prop_name As String, _
    ByVal long_border_specifier As Long, _
    ByVal dict_color_setting As Variant, _
    ByVal long_me_collor_no As Long)
    
    
    '結合セル同士でない場合のみチェックする
    If (range_me.MergeArea.Cells(1, 1).Address <> range_vs.MergeArea.Cells(1, 1).Address) Then
    
        variant_val2_of_vs = range_vs.MergeArea.Cells(1, 1).Value2 '結合セルの場合は左上セル値を使用
        str_vs_prop_name = "prop_" & TypeName(variant_val2_of_vs) & "_" & CStr(variant_val2_of_vs)
        
        'Color Dictionaryに登録している vs セルのみ評価する
        If ( _
            (str_me_prop_name <> str_vs_prop_name) And _
            (dict_color_setting.Exists(str_vs_prop_name)) _
        ) Then
            
            'カラー設定が同じ場合
            If (dict_color_setting.Item(str_vs_prop_name) = long_me_collor_no) Then
                With range_me.Borders(long_border_specifier)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
            End If
            
        End If
        
    End If
    
End Function

