# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version 4.2.1-0-g80c4cb6)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import wx.dataview
import wx.grid

import gettext
_ = gettext.gettext

###########################################################################
## Class CostCheckerMainFrame
###########################################################################

class CostCheckerMainFrame ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 1000,1000 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        self.m_menubar1 = wx.MenuBar( 0 )
        self.file_menu = wx.Menu()
        self.m_menuItem1 = wx.MenuItem( self.file_menu, wx.ID_ANY, _(u"新建项目"), _(u"新建一次测试内容导入项目"), wx.ITEM_NORMAL )
        self.file_menu.Append( self.m_menuItem1 )

        self.m_menuItem2 = wx.MenuItem( self.file_menu, wx.ID_ANY, _(u"读取记录"), _(u"读取本次的记录"), wx.ITEM_NORMAL )
        self.file_menu.Append( self.m_menuItem2 )

        self.m_menuItem3 = wx.MenuItem( self.file_menu, wx.ID_ANY, _(u"退出"), _(u"退出该软件"), wx.ITEM_NORMAL )
        self.file_menu.Append( self.m_menuItem3 )

        self.m_menubar1.Append( self.file_menu, _(u"文件") )

        self.edit_menu = wx.Menu()
        self.m_menubar1.Append( self.edit_menu, _(u"编辑") )

        self.SetMenuBar( self.m_menubar1 )

        self.m_statusBar1 = self.CreateStatusBar( 1, wx.STB_SIZEGRIP, wx.ID_ANY )
        main_sizer = wx.BoxSizer( wx.VERTICAL )

        titlebar_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, _(u"普通话测试站") ), wx.VERTICAL )

        self.m_staticText1 = wx.StaticText( titlebar_sizer.GetStaticBox(), wx.ID_ANY, _(u"山西大学普通话测试站 - 缴费名单确认软件 - Ver 2026.04"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL|wx.ST_ELLIPSIZE_MIDDLE|wx.ST_NO_AUTORESIZE )
        self.m_staticText1.Wrap( -1 )

        self.m_staticText1.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_MENU ) )
        self.m_staticText1.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )

        titlebar_sizer.Add( self.m_staticText1, 0, wx.ALL|wx.EXPAND, 5 )


        main_sizer.Add( titlebar_sizer, 0, wx.EXPAND, 5 )

        self.m_notebook1 = wx.Notebook( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_panel1 = wx.Panel( self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        bSizer10 = wx.BoxSizer( wx.VERTICAL )

        file_select_sizer = wx.BoxSizer( wx.HORIZONTAL )

        total_chart_select_sizer = wx.StaticBoxSizer( wx.StaticBox( self.m_panel1, wx.ID_ANY, _(u"受测学生总表") ), wx.VERTICAL )

        self.m_totlechart_filePicker1 = wx.FilePickerCtrl( total_chart_select_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, _(u"选择该学院报名总表"), _(u"Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"), wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
        total_chart_select_sizer.Add( self.m_totlechart_filePicker1, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_totalchart_picker_staticText2 = wx.StaticText( total_chart_select_sizer.GetStaticBox(), wx.ID_ANY, _(u"等待选择文件"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_totalchart_picker_staticText2.Wrap( -1 )

        total_chart_select_sizer.Add( self.m_totalchart_picker_staticText2, 0, wx.ALL|wx.EXPAND, 5 )


        file_select_sizer.Add( total_chart_select_sizer, 1, wx.ALIGN_CENTER, 5 )

        cost_chart_select_sizer = wx.StaticBoxSizer( wx.StaticBox( self.m_panel1, wx.ID_ANY, _(u"核对表目") ), wx.VERTICAL )

        needcost_picker_bSizer3 = wx.BoxSizer( wx.HORIZONTAL )

        self.m_staticText6 = wx.StaticText( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, _(u"需缴费学生名单"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText6.Wrap( -1 )

        needcost_picker_bSizer3.Add( self.m_staticText6, 0, wx.ALL, 5 )

        self.m_needcost_chart_filePicker11 = wx.FilePickerCtrl( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, _(u"选择该学院需缴费名单"), _(u"Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"), wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
        needcost_picker_bSizer3.Add( self.m_needcost_chart_filePicker11, 1, 0, 5 )


        cost_chart_select_sizer.Add( needcost_picker_bSizer3, 1, wx.EXPAND, 5 )

        self.m_needcost_picker_staticText21 = wx.StaticText( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, _(u"等待选择文件"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_needcost_picker_staticText21.Wrap( -1 )

        cost_chart_select_sizer.Add( self.m_needcost_picker_staticText21, 0, wx.ALL|wx.EXPAND, 5 )

        noneed_picker_bSizer31 = wx.BoxSizer( wx.HORIZONTAL )

        self.m_staticText61 = wx.StaticText( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, _(u"免缴费学生名单"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText61.Wrap( -1 )

        noneed_picker_bSizer31.Add( self.m_staticText61, 0, wx.ALL, 5 )

        self.m_noneed_chart_filePicker111 = wx.FilePickerCtrl( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, _(u"选择该学院需缴费名单"), _(u"Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"), wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
        noneed_picker_bSizer31.Add( self.m_noneed_chart_filePicker111, 1, 0, 5 )


        cost_chart_select_sizer.Add( noneed_picker_bSizer31, 1, wx.EXPAND, 5 )

        self.m_noneed_picker_staticText211 = wx.StaticText( cost_chart_select_sizer.GetStaticBox(), wx.ID_ANY, _(u"等待选择文件"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_noneed_picker_staticText211.Wrap( -1 )

        cost_chart_select_sizer.Add( self.m_noneed_picker_staticText211, 0, wx.ALL|wx.EXPAND, 5 )


        file_select_sizer.Add( cost_chart_select_sizer, 1, wx.EXPAND, 5 )


        bSizer10.Add( file_select_sizer, 0, wx.EXPAND, 5 )

        result_bSizer7 = wx.BoxSizer( wx.VERTICAL )

        short_result_bSizer8 = wx.BoxSizer( wx.HORIZONTAL )

        self.short_result_text = wx.StaticText( self.m_panel1, wx.ID_ANY, _(u"【检查结果】总人数：000\t\t需缴费学生人数：000\t\t免缴费学生人数：000\t\t错误数据：000\t错误占比：000%"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.short_result_text.Wrap( -1 )

        short_result_bSizer8.Add( self.short_result_text, 1, wx.ALL, 5 )


        result_bSizer7.Add( short_result_bSizer8, 0, wx.EXPAND, 5 )

        self.m_read_result_textCtrl1 = wx.TextCtrl( self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_RICH2 )
        result_bSizer7.Add( self.m_read_result_textCtrl1, 1, wx.ALL|wx.EXPAND, 5 )

        bSizer9 = wx.BoxSizer( wx.HORIZONTAL )

        self.m_staticText15 = wx.StaticText( self.m_panel1, wx.ID_ANY, _(u"错误数据会集中到第二页面，请切换到第二页面中更改"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText15.Wrap( -1 )

        bSizer9.Add( self.m_staticText15, 2, wx.ALL, 5 )

        self.m_staticText16 = wx.StaticText( self.m_panel1, wx.ID_ANY, _(u"当前数据："), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText16.Wrap( -1 )

        bSizer9.Add( self.m_staticText16, 0, wx.ALL, 5 )

        m_school_choice1Choices = [ _(u"文学院"), _(u"历史文化学院"), _(u"哲学学院"), _(u"外国语学院"), _(u"教育科学学院"), _(u"初民学院"), _(u"马克思主义学院"), _(u"新闻学院"), _(u"考古文博学院"), _(u"数学与统计学院"), _(u"计算机与信息技术学院"), _(u"物理电子工程学院"), _(u"化学化工学院"), _(u"体育学院"), _(u"音乐学院"), _(u"美术学院"), _(u"继续教育学院"), _(u"国际教育交流学院"), _(u"政治与公共管理学院"), _(u"法学院"), _(u"经济与管理学院"), _(u"生命科学学院"), _(u"环境与资源学院"), _(u"电力与建筑学院"), _(u"自动化与软件学院") ]
        self.m_school_choice1 = wx.Choice( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_school_choice1Choices, 0 )
        self.m_school_choice1.SetSelection( 0 )
        bSizer9.Add( self.m_school_choice1, 1, wx.ALL, 5 )

        self.m_button1 = wx.Button( self.m_panel1, wx.ID_ANY, _(u"保存并切换"), wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer9.Add( self.m_button1, 0, 0, 5 )


        result_bSizer7.Add( bSizer9, 0, wx.EXPAND, 5 )


        bSizer10.Add( result_bSizer7, 1, wx.EXPAND, 5 )


        self.m_panel1.SetSizer( bSizer10 )
        self.m_panel1.Layout()
        bSizer10.Fit( self.m_panel1 )
        self.m_notebook1.AddPage( self.m_panel1, _(u"第一页面-导入"), True )
        self.m_panel2 = wx.Panel( self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        bSizer12 = wx.BoxSizer( wx.VERTICAL )

        self.m_notebook2 = wx.Notebook( self.m_panel2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.NB_LEFT|wx.NB_NOPAGETHEME )
        self.m_notebook2.SetFont( wx.Font( wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString ) )

        self.m_panel4 = wx.Panel( self.m_notebook2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        bSizer11 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText17 = wx.StaticText( self.m_panel4, wx.ID_ANY, _(u"导入的错误数据"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText17.Wrap( -1 )

        bSizer11.Add( self.m_staticText17, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_dataViewListCtrl1 = wx.dataview.DataViewListCtrl( self.m_panel4, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.dataview.DV_HORIZ_RULES|wx.dataview.DV_ROW_LINES|wx.dataview.DV_VERT_RULES )
        self.m_schoolnumber_dataViewListColumn1 = self.m_dataViewListCtrl1.AppendTextColumn( _(u"学号"), wx.dataview.DATAVIEW_CELL_EDITABLE, 150, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_name_dataViewListColumn2 = self.m_dataViewListCtrl1.AppendTextColumn( _(u"姓名"), wx.dataview.DATAVIEW_CELL_EDITABLE, 100, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_idcard_dataViewListColumn3 = self.m_dataViewListCtrl1.AppendTextColumn( _(u"身份证信息"), wx.dataview.DATAVIEW_CELL_EDITABLE, 260, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_costneed_dataViewListColumn4 = self.m_dataViewListCtrl1.AppendToggleColumn( _(u"缴费需否"), wx.dataview.DATAVIEW_CELL_EDITABLE, 100, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_wrongreason_dataViewListColumn5 = self.m_dataViewListCtrl1.AppendTextColumn( _(u"数据错误原因"), wx.dataview.DATAVIEW_CELL_INERT, 270, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        bSizer11.Add( self.m_dataViewListCtrl1, 1, wx.ALL|wx.EXPAND, 5 )


        self.m_panel4.SetSizer( bSizer11 )
        self.m_panel4.Layout()
        bSizer11.Fit( self.m_panel4 )
        self.m_notebook2.AddPage( self.m_panel4, _(u"有误的数据"), True )
        self.m_panel5 = wx.Panel( self.m_notebook2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        bSizer111 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText171 = wx.StaticText( self.m_panel5, wx.ID_ANY, _(u"该学院的全部数据"), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText171.Wrap( -1 )

        bSizer111.Add( self.m_staticText171, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_grid11 = wx.grid.Grid( self.m_panel5, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )

        # Grid
        self.m_grid11.CreateGrid( 5, 5 )
        self.m_grid11.EnableEditing( True )
        self.m_grid11.EnableGridLines( True )
        self.m_grid11.EnableDragGridSize( True )
        self.m_grid11.SetMargins( 0, 0 )

        # Columns
        self.m_grid11.SetColSize( 0, 130 )
        self.m_grid11.SetColSize( 1, 140 )
        self.m_grid11.SetColSize( 2, 250 )
        self.m_grid11.SetColSize( 3, 90 )
        self.m_grid11.SetColSize( 4, 90 )
        self.m_grid11.EnableDragColMove( False )
        self.m_grid11.EnableDragColSize( True )
        self.m_grid11.SetColLabelValue( 0, _(u"学号") )
        self.m_grid11.SetColLabelValue( 1, _(u"姓名") )
        self.m_grid11.SetColLabelValue( 2, _(u"身分证信息") )
        self.m_grid11.SetColLabelValue( 3, _(u"缴费情况") )
        self.m_grid11.SetColLabelValue( 4, _(u"毕业年级") )
        self.m_grid11.SetColLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

        # Rows
        self.m_grid11.EnableDragRowSize( True )
        self.m_grid11.SetRowLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

        # Label Appearance

        # Cell Defaults
        self.m_grid11.SetDefaultCellAlignment( wx.ALIGN_LEFT, wx.ALIGN_TOP )
        bSizer111.Add( self.m_grid11, 1, wx.ALL|wx.EXPAND, 5 )


        self.m_panel5.SetSizer( bSizer111 )
        self.m_panel5.Layout()
        bSizer111.Fit( self.m_panel5 )
        self.m_notebook2.AddPage( self.m_panel5, _(u"全部导入内容"), False )

        bSizer12.Add( self.m_notebook2, 1, wx.EXPAND |wx.ALL, 5 )

        bSizer91 = wx.BoxSizer( wx.HORIZONTAL )

        self.m_staticText151 = wx.StaticText( self.m_panel2, wx.ID_ANY, _(u"修改的内容会加粗表示，修改后请保存数据"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText151.Wrap( -1 )

        bSizer91.Add( self.m_staticText151, 2, wx.ALL, 5 )

        self.m_staticText161 = wx.StaticText( self.m_panel2, wx.ID_ANY, _(u"当前查看："), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText161.Wrap( -1 )

        bSizer91.Add( self.m_staticText161, 0, wx.ALL, 5 )

        m_school_choice2Choices = [ _(u"文学院"), _(u"历史文化学院"), _(u"哲学学院"), _(u"外国语学院"), _(u"教育科学学院"), _(u"初民学院"), _(u"马克思主义学院"), _(u"新闻学院"), _(u"考古文博学院"), _(u"数学与统计学院"), _(u"计算机与信息技术学院"), _(u"物理电子工程学院"), _(u"化学化工学院"), _(u"体育学院"), _(u"音乐学院"), _(u"美术学院"), _(u"继续教育学院"), _(u"国际教育交流学院"), _(u"政治与公共管理学院"), _(u"法学院"), _(u"经济与管理学院"), _(u"生命科学学院"), _(u"环境与资源学院"), _(u"电力与建筑学院"), _(u"自动化与软件学院") ]
        self.m_school_choice2 = wx.Choice( self.m_panel2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_school_choice2Choices, 0 )
        self.m_school_choice2.SetSelection( 0 )
        bSizer91.Add( self.m_school_choice2, 1, wx.ALL, 5 )

        self.m_save_change_button11 = wx.Button( self.m_panel2, wx.ID_ANY, _(u"保存修改"), wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer91.Add( self.m_save_change_button11, 0, wx.ALL, 5 )


        bSizer12.Add( bSizer91, 0, wx.EXPAND, 5 )


        self.m_panel2.SetSizer( bSizer12 )
        self.m_panel2.Layout()
        bSizer12.Fit( self.m_panel2 )
        self.m_notebook1.AddPage( self.m_panel2, _(u"第二页面-核对"), False )
        self.m_panel3 = wx.Panel( self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        bSizer13 = wx.BoxSizer( wx.VERTICAL )

        self.m_dataViewListCtrl2 = wx.dataview.DataViewListCtrl( self.m_panel3, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.dataview.DV_HORIZ_RULES|wx.dataview.DV_ROW_LINES|wx.dataview.DV_VERT_RULES )
        self.m_schoolname_dataViewListColumn6 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"学院"), wx.dataview.DATAVIEW_CELL_INERT, 250, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_checkin_dataViewListColumn7 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"报名人数"), wx.dataview.DATAVIEW_CELL_INERT, 120, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_needpay_dataViewListColumn8 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"应缴人数"), wx.dataview.DATAVIEW_CELL_INERT, 120, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_nocost_dataViewListColumn9 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"免缴费数"), wx.dataview.DATAVIEW_CELL_INERT, 120, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_unfix_dataViewListColumn11 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"未修正错误"), wx.dataview.DATAVIEW_CELL_INERT, 120, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        self.m_shouldpaymoney_dataViewListColumn10 = self.m_dataViewListCtrl2.AppendTextColumn( _(u"应缴费款"), wx.dataview.DATAVIEW_CELL_INERT, 120, wx.ALIGN_LEFT, wx.dataview.DATAVIEW_COL_RESIZABLE )
        bSizer13.Add( self.m_dataViewListCtrl2, 3, wx.ALL|wx.EXPAND, 5 )

        self.m_totalexport_staticText14 = wx.StaticText( self.m_panel3, wx.ID_ANY, _(u"总报名人数：000人\t\t坞城校区：000人\t\t东山校区：000人"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_totalexport_staticText14.Wrap( -1 )

        bSizer13.Add( self.m_totalexport_staticText14, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_wucheng_stats_staticText152 = wx.StaticText( self.m_panel3, wx.ID_ANY, _(u"\t其中，坞城校区——\t非毕业年级：000人\t\t毕业年级：000人"), wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_wucheng_stats_staticText152.Wrap( -1 )

        bSizer13.Add( self.m_wucheng_stats_staticText152, 0, wx.ALL|wx.EXPAND, 5 )

        bSizer14 = wx.BoxSizer( wx.HORIZONTAL )

        self.m_staticText162 = wx.StaticText( self.m_panel3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText162.Wrap( -1 )

        bSizer14.Add( self.m_staticText162, 1, wx.ALL, 5 )

        self.m_export_splitschool_button3 = wx.Button( self.m_panel3, wx.ID_ANY, _(u"分学院导出表格"), wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer14.Add( self.m_export_splitschool_button3, 0, wx.ALL, 5 )

        self.m_export_thewhole_button4 = wx.Button( self.m_panel3, wx.ID_ANY, _(u"导出总表"), wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer14.Add( self.m_export_thewhole_button4, 0, wx.ALL, 5 )


        bSizer13.Add( bSizer14, 0, wx.EXPAND, 5 )


        self.m_panel3.SetSizer( bSizer13 )
        self.m_panel3.Layout()
        bSizer13.Fit( self.m_panel3 )
        self.m_notebook1.AddPage( self.m_panel3, _(u"第三页面-导出"), False )

        main_sizer.Add( self.m_notebook1, 1, wx.EXPAND |wx.ALL, 5 )


        self.SetSizer( main_sizer )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.Bind( wx.EVT_MENU, self.on_new_project_button_click, id = self.m_menuItem1.GetId() )
        self.Bind( wx.EVT_MENU, self.on_read_record_button_click, id = self.m_menuItem2.GetId() )
        self.Bind( wx.EVT_MENU, self.on_exit_button_click, id = self.m_menuItem3.GetId() )
        self.m_staticText1.Bind( wx.EVT_LEFT_DCLICK, self.on_yanlun_double_click )
        self.m_totlechart_filePicker1.Bind( wx.EVT_FILEPICKER_CHANGED, self.on_all_tocheck_chart_file_change )
        self.m_needcost_chart_filePicker11.Bind( wx.EVT_FILEPICKER_CHANGED, self.on_needpay_chart_file_change )
        self.m_noneed_chart_filePicker111.Bind( wx.EVT_FILEPICKER_CHANGED, self.on_nocost_chart_file_change )
        self.m_school_choice1.Bind( wx.EVT_CHOICE, self.on_read_data_school_change )
        self.m_button1.Bind( wx.EVT_BUTTON, self.on_read_data_save_button_click )
        self.m_dataViewListCtrl1.Bind( wx.dataview.EVT_DATAVIEW_ITEM_EDITING_DONE, self.on_wrong_data_item_editing_done, id = wx.ID_ANY )
        self.m_dataViewListCtrl1.Bind( wx.dataview.EVT_DATAVIEW_ITEM_START_EDITING, self.on_wrong_data_item_editing_start, id = wx.ID_ANY )
        self.m_dataViewListCtrl1.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.on_wrong_data_item_value_changed, id = wx.ID_ANY )
        self.m_grid11.Bind( wx.grid.EVT_GRID_CELL_CHANGED, self.on_grid_cell_change )
        self.m_grid11.Bind( wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.on_grid_cell_left_click )
        self.m_grid11.Bind( wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_grid_cell_right_click )
        self.m_export_splitschool_button3.Bind( wx.EVT_BUTTON, self.on_export_each_school_button_click )
        self.m_export_thewhole_button4.Bind( wx.EVT_BUTTON, self.on_export_total_chart_button_click )

    def __del__( self ):
        # Disconnect Events
        self.Unbind( wx.EVT_MENU, id = self.m_menuItem1.GetId() )
        self.Unbind( wx.EVT_MENU, id = self.m_menuItem2.GetId() )
        self.Unbind( wx.EVT_MENU, id = self.m_menuItem3.GetId() )
        self.m_staticText1.Unbind( wx.EVT_LEFT_DCLICK, None )
        self.m_totlechart_filePicker1.Unbind( wx.EVT_FILEPICKER_CHANGED, None )
        self.m_needcost_chart_filePicker11.Unbind( wx.EVT_FILEPICKER_CHANGED, None )
        self.m_noneed_chart_filePicker111.Unbind( wx.EVT_FILEPICKER_CHANGED, None )
        self.m_school_choice1.Unbind( wx.EVT_CHOICE, None )
        self.m_button1.Unbind( wx.EVT_BUTTON, None )
        self.m_dataViewListCtrl1.Unbind( wx.dataview.EVT_DATAVIEW_ITEM_EDITING_DONE, None, id = wx.ID_ANY )
        self.m_dataViewListCtrl1.Unbind( wx.dataview.EVT_DATAVIEW_ITEM_START_EDITING, None, id = wx.ID_ANY )
        self.m_dataViewListCtrl1.Unbind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, None, id = wx.ID_ANY )
        self.m_grid11.Unbind( wx.grid.EVT_GRID_CELL_CHANGED, None )
        self.m_grid11.Unbind( wx.grid.EVT_GRID_CELL_LEFT_CLICK, None )
        self.m_grid11.Unbind( wx.grid.EVT_GRID_CELL_RIGHT_CLICK, None )
        self.m_export_splitschool_button3.Unbind( wx.EVT_BUTTON, None )
        self.m_export_thewhole_button4.Unbind( wx.EVT_BUTTON, None )


    # Virtual event handlers, override them in your derived class
    def on_new_project_button_click( self, event ):
        event.Skip()

    def on_read_record_button_click( self, event ):
        event.Skip()

    def on_exit_button_click( self, event ):
        event.Skip()

    def on_yanlun_double_click( self, event ):
        event.Skip()

    def on_all_tocheck_chart_file_change( self, event ):
        event.Skip()

    def on_needpay_chart_file_change( self, event ):
        event.Skip()

    def on_nocost_chart_file_change( self, event ):
        event.Skip()

    def on_read_data_school_change( self, event ):
        event.Skip()

    def on_read_data_save_button_click( self, event ):
        event.Skip()

    def on_wrong_data_item_editing_done( self, event ):
        event.Skip()

    def on_wrong_data_item_editing_start( self, event ):
        event.Skip()

    def on_wrong_data_item_value_changed( self, event ):
        event.Skip()

    def on_grid_cell_change( self, event ):
        event.Skip()

    def on_grid_cell_left_click( self, event ):
        event.Skip()

    def on_grid_cell_right_click( self, event ):
        event.Skip()

    def on_export_each_school_button_click( self, event ):
        event.Skip()

    def on_export_total_chart_button_click( self, event ):
        event.Skip()

    # Virtual image path resolution method. Override this in your derived class.
    def img_wrapper( self, bitmap_path ):
        return bitmap_path


