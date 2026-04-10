# -*- coding: utf-8 -*-


import wx
import wx.xrc
import wx.dataview
import wx.grid

import json
import gettext

from datetime import datetime, date
from typing import Dict, List, Optional, Union, Set
from rich import print as prt
from pathlib import Path

from yanlun import YanlunText
from chart_read import PscPersonManager, CheckPerson

_ = gettext.gettext

__version__ = "2026.04"

SCHOOL_NAMES = [
    "文学院",
    "历史文化学院",
    "哲学学院",
    "外国语学院",
    "教育科学学院",
    "初民学院",
    "马克思主义学院",
    "新闻学院",
    "考古文博学院",
    "数学与统计学院",
    "计算机与信息技术学院",
    "物理电子工程学院",
    "化学化工学院",
    "体育学院",
    "音乐学院",
    "美术学院",
    "继续教育学院",
    "国际教育交流学院",
    "政治与公共管理学院",
    "法学院",
    "经济与管理学院",
    "生命科学学院",
    "环境与资源学院",
    "电力与建筑学院",
    "自动化与软件学院",
]

SPERATE_SCHOOL_ZONE = {
    "坞城": [
        "文学院",
        "历史文化学院",
        "哲学学院",
        "外国语学院",
        "教育科学学院",
        "初民学院",
        "马克思主义学院",
        "新闻学院",
        "考古文博学院",
        "数学与统计学院",
        "计算机与信息技术学院",
        "物理电子工程学院",
        "化学化工学院",
        "体育学院",
        "音乐学院",
        "美术学院",
        "继续教育学院",
        "国际教育交流学院",
    ],
    "东山": [
        "政治与公共管理学院",
        "法学院",
        "经济与管理学院",
        "生命科学学院",
        "环境与资源学院",
        "电力与建筑学院",
        "自动化与软件学院",
    ],
}


RESULT_PATH = Path("./result").resolve()

RESULT_PATH.mkdir(parents=True, exist_ok=True)


###########################################################################
## Class CostCheckerMainFrame
###########################################################################


class CostCheckerMainFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title=wx.EmptyString,
            pos=wx.DefaultPosition,
            size=wx.Size(1000, 1000),
            style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL,
        )

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        self.m_menubar1 = wx.MenuBar(0)
        self.file_menu = wx.Menu()
        self.m_menuItem1 = wx.MenuItem(
            self.file_menu,
            wx.ID_ANY,
            _("新建项目"),
            _("新建一次测试内容导入项目"),
            wx.ITEM_NORMAL,
        )
        self.file_menu.Append(self.m_menuItem1)

        self.m_menuItem2 = wx.MenuItem(
            self.file_menu,
            wx.ID_ANY,
            _("读取记录"),
            _("读取本次的记录"),
            wx.ITEM_NORMAL,
        )
        self.file_menu.Append(self.m_menuItem2)

        self.m_menuItem3 = wx.MenuItem(
            self.file_menu, wx.ID_ANY, _("退出"), _("退出该软件"), wx.ITEM_NORMAL
        )
        self.file_menu.Append(self.m_menuItem3)

        self.m_menubar1.Append(self.file_menu, _("文件"))

        self.edit_menu = wx.Menu()
        self.m_menubar1.Append(self.edit_menu, _("编辑"))

        self.SetMenuBar(self.m_menubar1)

        self.m_statusBar1 = self.CreateStatusBar(1, wx.STB_SIZEGRIP, wx.ID_ANY)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        titlebar_sizer = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, _("普通话测试站")), wx.VERTICAL
        )

        self.m_yanluntext1 = YanlunText(
            titlebar_sizer.GetStaticBox(),
            wx.ID_ANY,
            __version__,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_CENTER_HORIZONTAL | wx.ST_ELLIPSIZE_MIDDLE | wx.ST_NO_AUTORESIZE,
        )

        titlebar_sizer.Add(self.m_yanluntext1, 0, wx.ALL | wx.EXPAND, 5)

        main_sizer.Add(titlebar_sizer, 0, wx.EXPAND, 5)

        self.m_notebook1 = wx.Notebook(
            self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0
        )
        self.m_panel1 = wx.Panel(
            self.m_notebook1,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TAB_TRAVERSAL,
        )
        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        file_select_sizer = wx.BoxSizer(wx.HORIZONTAL)

        total_chart_select_sizer = wx.StaticBoxSizer(
            wx.StaticBox(self.m_panel1, wx.ID_ANY, _("受测学生总表")), wx.VERTICAL
        )

        self.m_totalchart_filePicker1 = wx.FilePickerCtrl(
            total_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            wx.EmptyString,
            _("选择该学院报名总表"),
            _("Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.FLP_DEFAULT_STYLE,
        )
        total_chart_select_sizer.Add(
            self.m_totalchart_filePicker1, 0, wx.ALL | wx.EXPAND, 5
        )

        self.m_totalchart_picker_staticText2 = wx.StaticText(
            total_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            _("等待选择文件"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_RIGHT,
        )
        self.m_totalchart_picker_staticText2.Wrap(-1)

        total_chart_select_sizer.Add(
            self.m_totalchart_picker_staticText2, 0, wx.ALL | wx.EXPAND, 5
        )

        file_select_sizer.Add(total_chart_select_sizer, 1, wx.ALIGN_CENTER, 5)

        cost_chart_select_sizer = wx.StaticBoxSizer(
            wx.StaticBox(self.m_panel1, wx.ID_ANY, _("核对表目")), wx.VERTICAL
        )

        needcost_picker_bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText6 = wx.StaticText(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            _("需缴费学生名单"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText6.Wrap(-1)

        needcost_picker_bSizer3.Add(self.m_staticText6, 0, wx.ALL, 5)

        self.m_needcost_chart_filePicker11 = wx.FilePickerCtrl(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            wx.EmptyString,
            _("选择该学院需缴费名单"),
            _("Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.FLP_DEFAULT_STYLE,
        )
        needcost_picker_bSizer3.Add(self.m_needcost_chart_filePicker11, 1, 0, 5)

        cost_chart_select_sizer.Add(needcost_picker_bSizer3, 1, wx.EXPAND, 5)

        self.m_needcost_picker_staticText21 = wx.StaticText(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            _("等待选择文件"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_RIGHT,
        )
        self.m_needcost_picker_staticText21.Wrap(-1)

        cost_chart_select_sizer.Add(
            self.m_needcost_picker_staticText21, 0, wx.ALL | wx.EXPAND, 5
        )

        noneed_picker_bSizer31 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText61 = wx.StaticText(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            _("免缴费学生名单"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText61.Wrap(-1)

        noneed_picker_bSizer31.Add(self.m_staticText61, 0, wx.ALL, 5)

        self.m_noneed_chart_filePicker111 = wx.FilePickerCtrl(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            wx.EmptyString,
            _("选择该学院需缴费名单"),
            _("Excel 表格文件 (*.xlsx;*.xls)|*.xlsx;*.xls"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.FLP_DEFAULT_STYLE,
        )
        noneed_picker_bSizer31.Add(self.m_noneed_chart_filePicker111, 1, 0, 5)

        cost_chart_select_sizer.Add(noneed_picker_bSizer31, 1, wx.EXPAND, 5)

        self.m_noneed_picker_staticText211 = wx.StaticText(
            cost_chart_select_sizer.GetStaticBox(),
            wx.ID_ANY,
            _("等待选择文件"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_RIGHT,
        )
        self.m_noneed_picker_staticText211.Wrap(-1)

        cost_chart_select_sizer.Add(
            self.m_noneed_picker_staticText211, 0, wx.ALL | wx.EXPAND, 5
        )

        file_select_sizer.Add(cost_chart_select_sizer, 1, wx.EXPAND, 5)

        bSizer10.Add(file_select_sizer, 0, wx.EXPAND, 5)

        result_bSizer7 = wx.BoxSizer(wx.VERTICAL)

        short_result_bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.short_result_text = wx.StaticText(
            self.m_panel1,
            wx.ID_ANY,
            _(
                "【检查结果】总人数：000\t\t需缴费学生人数：000\t\t免缴费学生人数：000\t\t错误数据：000\t错误占比：000%"
            ),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.short_result_text.Wrap(-1)

        short_result_bSizer8.Add(self.short_result_text, 1, wx.ALL, 5)

        result_bSizer7.Add(short_result_bSizer8, 0, wx.EXPAND, 5)

        self.m_read_result_textCtrl1 = wx.TextCtrl(
            self.m_panel1,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
        )
        result_bSizer7.Add(self.m_read_result_textCtrl1, 1, wx.ALL | wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText15 = wx.StaticText(
            self.m_panel1,
            wx.ID_ANY,
            _("错误数据会集中到第二页面，请切换到第二页面中更改"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText15.Wrap(-1)

        bSizer9.Add(self.m_staticText15, 2, wx.ALL, 5)

        self.m_staticText16 = wx.StaticText(
            self.m_panel1,
            wx.ID_ANY,
            _("当前数据："),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText16.Wrap(-1)

        bSizer9.Add(self.m_staticText16, 0, wx.ALL, 5)

        m_school_choice1Choices = SCHOOL_NAMES.copy()
        self.m_school_choice1 = wx.Choice(
            self.m_panel1,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            m_school_choice1Choices,
            0,
        )
        self.m_school_choice1.SetSelection(0)
        bSizer9.Add(self.m_school_choice1, 1, wx.ALL, 5)

        self.m_button1 = wx.Button(
            self.m_panel1,
            wx.ID_ANY,
            _("保存并切换"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer9.Add(self.m_button1, 0, 0, 5)

        result_bSizer7.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer10.Add(result_bSizer7, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        self.m_notebook1.AddPage(self.m_panel1, _("第一页面-导入"), True)
        self.m_panel2 = wx.Panel(
            self.m_notebook1,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TAB_TRAVERSAL,
        )
        bSizer12 = wx.BoxSizer(wx.VERTICAL)

        self.m_notebook2 = wx.Notebook(
            self.m_panel2,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.NB_LEFT | wx.NB_NOPAGETHEME,
        )

        self.m_panel4 = wx.Panel(
            self.m_notebook2,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TAB_TRAVERSAL,
        )
        bSizer11 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText17 = wx.StaticText(
            self.m_panel4,
            wx.ID_ANY,
            _("导入的错误数据"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_CENTER_HORIZONTAL,
        )
        self.m_staticText17.Wrap(-1)

        bSizer11.Add(self.m_staticText17, 0, wx.ALL | wx.EXPAND, 5)

        self.m_dataViewListCtrl1 = wx.dataview.DataViewListCtrl(
            self.m_panel4,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.dataview.DV_HORIZ_RULES
            | wx.dataview.DV_ROW_LINES
            | wx.dataview.DV_VERT_RULES,
        )
        self.m_schoolnumber_dataViewListColumn1 = (
            self.m_dataViewListCtrl1.AppendTextColumn(
                _("学号"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                150,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_name_dataViewListColumn2 = self.m_dataViewListCtrl1.AppendTextColumn(
            _("姓名"),
            wx.dataview.DATAVIEW_CELL_EDITABLE,
            100,
            wx.ALIGN_LEFT,
            wx.dataview.DATAVIEW_COL_RESIZABLE,
        )
        self.m_idcard_dataViewListColumn3 = self.m_dataViewListCtrl1.AppendTextColumn(
            _("身份证信息"),
            wx.dataview.DATAVIEW_CELL_EDITABLE,
            260,
            wx.ALIGN_LEFT,
            wx.dataview.DATAVIEW_COL_RESIZABLE,
        )
        self.m_costneed_dataViewListColumn4 = (
            self.m_dataViewListCtrl1.AppendToggleColumn(
                _("缴费需否"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                100,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_wrongreason_dataViewListColumn5 = (
            self.m_dataViewListCtrl1.AppendTextColumn(
                _("数据错误原因"),
                wx.dataview.DATAVIEW_CELL_INERT,
                270,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        bSizer11.Add(self.m_dataViewListCtrl1, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel4.SetSizer(bSizer11)
        self.m_panel4.Layout()
        bSizer11.Fit(self.m_panel4)
        self.m_notebook2.AddPage(self.m_panel4, _("有误的数据"), True)
        self.m_panel5 = wx.Panel(
            self.m_notebook2,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TAB_TRAVERSAL,
        )
        bSizer111 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText171 = wx.StaticText(
            self.m_panel5,
            wx.ID_ANY,
            _("该学院的全部数据"),
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_CENTER_HORIZONTAL,
        )
        self.m_staticText171.Wrap(-1)

        bSizer111.Add(self.m_staticText171, 0, wx.ALL | wx.EXPAND, 5)

        self.m_grid11 = wx.grid.Grid(
            self.m_panel5, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0
        )

        # Grid
        self.m_grid11.CreateGrid(5, 5)
        self.m_grid11.EnableEditing(True)
        self.m_grid11.EnableGridLines(True)
        self.m_grid11.EnableDragGridSize(True)
        self.m_grid11.SetMargins(0, 0)

        # Columns
        self.m_grid11.SetColSize(0, 130)
        self.m_grid11.SetColSize(1, 140)
        self.m_grid11.SetColSize(2, 250)
        self.m_grid11.SetColSize(3, 90)
        self.m_grid11.SetColSize(4, 90)
        self.m_grid11.EnableDragColMove(False)
        self.m_grid11.EnableDragColSize(True)
        self.m_grid11.SetColLabelValue(0, _("学号"))
        self.m_grid11.SetColLabelValue(1, _("姓名"))
        self.m_grid11.SetColLabelValue(2, _("身分证信息"))
        self.m_grid11.SetColLabelValue(3, _("缴费情况"))
        self.m_grid11.SetColLabelValue(4, _("毕业年级"))
        self.m_grid11.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

        # Rows
        self.m_grid11.EnableDragRowSize(True)
        self.m_grid11.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

        # Label Appearance

        # Cell Defaults
        self.m_grid11.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
        bSizer111.Add(self.m_grid11, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel5.SetSizer(bSizer111)
        self.m_panel5.Layout()
        bSizer111.Fit(self.m_panel5)
        self.m_notebook2.AddPage(self.m_panel5, _("全部导入内容"), False)

        bSizer12.Add(self.m_notebook2, 1, wx.EXPAND | wx.ALL, 5)

        bSizer91 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText151 = wx.StaticText(
            self.m_panel2,
            wx.ID_ANY,
            _("修改的内容会加粗表示，修改后请保存数据"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText151.Wrap(-1)

        bSizer91.Add(self.m_staticText151, 2, wx.ALL, 5)

        self.m_staticText161 = wx.StaticText(
            self.m_panel2,
            wx.ID_ANY,
            _("当前查看："),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_staticText161.Wrap(-1)

        bSizer91.Add(self.m_staticText161, 0, wx.ALL, 5)

        m_school_choice2Choices = SCHOOL_NAMES.copy()
        self.m_school_choice2 = wx.Choice(
            self.m_panel2,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            m_school_choice2Choices,
            0,
        )
        self.m_school_choice2.SetSelection(0)
        bSizer91.Add(self.m_school_choice2, 1, wx.ALL, 5)

        self.m_save_change_button11 = wx.Button(
            self.m_panel2,
            wx.ID_ANY,
            _("保存修改"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer91.Add(self.m_save_change_button11, 0, wx.ALL, 5)

        bSizer12.Add(bSizer91, 0, wx.EXPAND, 5)

        self.m_panel2.SetSizer(bSizer12)
        self.m_panel2.Layout()
        bSizer12.Fit(self.m_panel2)
        self.m_notebook1.AddPage(self.m_panel2, _("第二页面-核对"), False)
        self.m_panel3 = wx.Panel(
            self.m_notebook1,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TAB_TRAVERSAL,
        )
        bSizer13 = wx.BoxSizer(wx.VERTICAL)

        self.m_export_total_dataViewListCtrl2 = wx.dataview.DataViewListCtrl(
            self.m_panel3,
            wx.ID_ANY,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.dataview.DV_HORIZ_RULES
            | wx.dataview.DV_ROW_LINES
            | wx.dataview.DV_VERT_RULES,
        )
        self.m_schoolname_dataViewListColumn6 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("学院"),
                wx.dataview.DATAVIEW_CELL_INERT,
                250,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_checkin_dataViewListColumn7 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("报名人数"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                120,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_needpay_dataViewListColumn8 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("应缴人数"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                120,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_nocost_dataViewListColumn9 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("免缴费数"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                120,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_unfix_dataViewListColumn11 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("未修正错误"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                120,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        self.m_shouldpaymoney_dataViewListColumn10 = (
            self.m_export_total_dataViewListCtrl2.AppendTextColumn(
                _("应缴费款"),
                wx.dataview.DATAVIEW_CELL_EDITABLE,
                120,
                wx.ALIGN_LEFT,
                wx.dataview.DATAVIEW_COL_RESIZABLE,
            )
        )
        bSizer13.Add(self.m_export_total_dataViewListCtrl2, 3, wx.ALL | wx.EXPAND, 5)

        self.m_totalexport_staticText14 = wx.StaticText(
            self.m_panel3,
            wx.ID_ANY,
            _("总报名人数：000人\t\t坞城校区：000人\t\t东山校区：000人"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_totalexport_staticText14.Wrap(-1)

        bSizer13.Add(self.m_totalexport_staticText14, 0, wx.ALL | wx.EXPAND, 5)

        self.m_wucheng_stats_staticText152 = wx.StaticText(
            self.m_panel3,
            wx.ID_ANY,
            _("\t其中，坞城校区——\t非毕业年级：000人\t\t毕业年级：000人"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.m_wucheng_stats_staticText152.Wrap(-1)

        bSizer13.Add(self.m_wucheng_stats_staticText152, 0, wx.ALL | wx.EXPAND, 5)

        bSizer14 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText162 = wx.StaticText(
            self.m_panel3,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.ALIGN_RIGHT,
        )
        self.m_staticText162.Wrap(-1)

        bSizer14.Add(self.m_staticText162, 1, wx.ALL, 5)

        self.m_export_splitschool_button3 = wx.Button(
            self.m_panel3,
            wx.ID_ANY,
            _("分学院导出表格"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer14.Add(self.m_export_splitschool_button3, 0, wx.ALL, 5)

        self.m_export_thewhole_button4 = wx.Button(
            self.m_panel3,
            wx.ID_ANY,
            _("导出总表"),
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        bSizer14.Add(self.m_export_thewhole_button4, 0, wx.ALL, 5)

        bSizer13.Add(bSizer14, 0, wx.EXPAND, 5)

        self.m_panel3.SetSizer(bSizer13)
        self.m_panel3.Layout()
        bSizer13.Fit(self.m_panel3)
        self.m_notebook1.AddPage(self.m_panel3, _("第三页面-导出"), False)

        main_sizer.Add(self.m_notebook1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(main_sizer)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.Bind(
            wx.EVT_MENU, self.on_new_project_button_click, id=self.m_menuItem1.GetId()
        )
        self.Bind(
            wx.EVT_MENU, self.on_read_record_button_click, id=self.m_menuItem2.GetId()
        )
        self.Bind(wx.EVT_MENU, self.on_exit_button_click, id=self.m_menuItem3.GetId())
        self.m_yanluntext1.Bind(wx.EVT_LEFT_DCLICK, self.on_yanlun_double_click)
        self.m_totalchart_filePicker1.Bind(
            wx.EVT_FILEPICKER_CHANGED, self.on_all_tocheck_chart_file_change
        )
        self.m_needcost_chart_filePicker11.Bind(
            wx.EVT_FILEPICKER_CHANGED, self.on_needpay_chart_file_change
        )
        self.m_noneed_chart_filePicker111.Bind(
            wx.EVT_FILEPICKER_CHANGED, self.on_nocost_chart_file_change
        )
        self.m_school_choice1.Bind(wx.EVT_CHOICE, self.on_read_data_school_change)
        self.m_button1.Bind(wx.EVT_BUTTON, self.on_read_data_save_button_click)
        self.m_notebook1.Bind(
            wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_page1_notebook_changed
        )
        self.m_notebook2.Bind(
            wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_page2_notebook_changed
        )
        self.m_dataViewListCtrl1.Bind(
            wx.dataview.EVT_DATAVIEW_ITEM_EDITING_DONE,
            self.on_wrong_data_item_editing_done,
            id=wx.ID_ANY,
        )
        self.m_dataViewListCtrl1.Bind(
            wx.dataview.EVT_DATAVIEW_ITEM_START_EDITING,
            self.on_wrong_data_item_editing_start,
            id=wx.ID_ANY,
        )
        self.m_dataViewListCtrl1.Bind(
            wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED,
            self.on_wrong_data_item_value_changed,
            id=wx.ID_ANY,
        )
        self.m_grid11.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_grid_cell_change)
        self.m_grid11.Bind(
            wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.on_grid_cell_left_click
        )
        self.m_grid11.Bind(
            wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_grid_cell_right_click
        )
        self.m_export_splitschool_button3.Bind(
            wx.EVT_BUTTON, self.on_export_each_school_button_click
        )
        self.m_export_thewhole_button4.Bind(
            wx.EVT_BUTTON, self.on_export_total_chart_button_click
        )

        # 初始化当前数据
        self.reset_whole_tocheck_chart()
        self.sp_school_pscmanagers = {schn: PscPersonManager() for schn in SCHOOL_NAMES}
        self.sp_school_errors = {schn: {} for schn in SCHOOL_NAMES}
        self.sp_school_stats = {schn: {} for schn in SCHOOL_NAMES}
        self.store_path: Optional[Path] = None

    def __del__(self):
        # Disconnect Events
        self.Unbind(wx.EVT_MENU, id=self.m_menuItem1.GetId())
        self.Unbind(wx.EVT_MENU, id=self.m_menuItem2.GetId())
        self.Unbind(wx.EVT_MENU, id=self.m_menuItem3.GetId())
        self.m_yanluntext1.Unbind(wx.EVT_LEFT_DCLICK, None)
        self.m_totalchart_filePicker1.Unbind(wx.EVT_FILEPICKER_CHANGED, None)
        self.m_needcost_chart_filePicker11.Unbind(wx.EVT_FILEPICKER_CHANGED, None)
        self.m_noneed_chart_filePicker111.Unbind(wx.EVT_FILEPICKER_CHANGED, None)
        self.m_school_choice1.Unbind(wx.EVT_CHOICE, None)
        self.m_button1.Unbind(wx.EVT_BUTTON, None)
        self.m_dataViewListCtrl1.Unbind(
            wx.dataview.EVT_DATAVIEW_ITEM_EDITING_DONE, None, id=wx.ID_ANY
        )
        self.m_dataViewListCtrl1.Unbind(
            wx.dataview.EVT_DATAVIEW_ITEM_START_EDITING, None, id=wx.ID_ANY
        )
        self.m_dataViewListCtrl1.Unbind(
            wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, None, id=wx.ID_ANY
        )
        self.m_grid11.Unbind(wx.grid.EVT_GRID_CELL_CHANGED, None)
        self.m_grid11.Unbind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, None)
        self.m_grid11.Unbind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, None)
        self.m_export_splitschool_button3.Unbind(wx.EVT_BUTTON, None)
        self.m_export_thewhole_button4.Unbind(wx.EVT_BUTTON, None)

    def initial_store_path(self):
        self.store_path = RESULT_PATH / datetime.now().strftime("%Y%m%d-%H%M%S")
        self.store_path.mkdir(parents=True, exist_ok=True)

    # Virtual event handlers, override them in your derived class
    def on_new_project_button_click(self, event):
        self.initial_store_path()

        self.sp_school_pscmanagers = {schn: PscPersonManager() for schn in SCHOOL_NAMES}
        self.sp_school_errors = {schn: {} for schn in SCHOOL_NAMES}
        self.sp_school_stats = {schn: {} for schn in SCHOOL_NAMES}
        self.reset_page1_all()

        event.Skip()

    def on_read_record_button_click(self, event):
        event.Skip()

    def on_exit_button_click(self, event):
        event.Skip()

    def on_yanlun_double_click(self, event):
        self.m_yanluntext1.SwitchYanlunLabel()
        event.Skip()

    def reset_whole_tocheck_chart(self):
        self.m_totalchart_filePicker1.SetPath("")
        self.m_totalchart_picker_staticText2.SetLabel("等待选择文件")
        self.current_pscmanager = PscPersonManager()
        self.current_psc_stats = {}
        self.current_psc_errors = {}

    def reset_needpay_chart(self):
        self.m_needcost_chart_filePicker11.SetPath("")
        self.m_needcost_picker_staticText21.SetLabel("等待选择文件")

    def reset_nocost_chart(self):
        self.m_noneed_chart_filePicker111.SetPath("")
        self.m_noneed_picker_staticText211.SetLabel("等待选择文件")

    def reset_read_result(self):
        self.short_result_text.SetLabelText(
            "【检查结果】总人数：000\t\t需缴费学生人数：000\t\t免缴费学生人数：000\t\t错误数据：000\t错误占比：000%"
        )
        self.m_read_result_textCtrl1.Clear()

    def reset_page1_all(self):
        self.reset_whole_tocheck_chart()
        self.reset_needpay_chart()
        self.reset_nocost_chart()
        self.reset_read_result()

    def on_all_tocheck_chart_file_change(self, event):
        self.m_totalchart_picker_staticText2.SetLabel(
            Path(self.m_totalchart_filePicker1.GetPath()).name
        )
        self.current_pscmanager: PscPersonManager = PscPersonManager.load_from_excel(
            self.m_totalchart_filePicker1.GetPath()
        )
        self.reset_needpay_chart()
        self.reset_nocost_chart()
        self.reset_read_result()
        # self.m_totlechart_filePicker1.GetPath()
        # event.Skip()

    def go_for_validate(self):
        self.current_psc_errors = {}
        for response in self.current_pscmanager.validate_against_fee_lists(
            exempt_file=self.m_noneed_chart_filePicker111.GetPath(),
            fee_file=self.m_needcost_chart_filePicker11.GetPath(),
        ):
            """
            {
                "总表项目序号": i,
                "学号": person.student_id,
                "姓名": person.name,
                "身份证": person.id_number,
                "需缴费": None,
                "原因": "身份证号格式错误",
            }
            """
            self.current_psc_errors[response["总表项目序号"]] = response
            self.m_read_result_textCtrl1.AppendText(
                f"{response['姓名']}{response['学号']}，证件号{response['身份证']}{'，无需缴费' if response['需缴费'] is False else ('，应缴费' if response['需缴费'] is True else '')}，{response['原因']}\n"
            )

        # {"total": 0, "shouldpay": 0, "exempt": 0, "error": 0, "grad": 0, "error_rate": 0.0"}
        self.current_psc_stats = self.current_pscmanager.get_summary_stats()
        self.short_result_text.SetLabelText(
            "【检查结果】总人数：{total:3d}\t\t需缴费学生人数：{shouldpay:3d}\t免缴费学生人数：{exempt:3d}\t错误数据：{error:3d}\t错误占比：{error_rate:.2f}%".format(
                **self.current_psc_stats
            )
        )
        self.current_pscmanager.get_summary_stats()

    def on_needpay_chart_file_change(self, event):
        # 先重置结果
        self.reset_read_result()

        # 改一下文件名提示
        self.m_needcost_picker_staticText21.SetLabel(
            Path(self.m_needcost_chart_filePicker11.GetPath()).name
        )

        # 是不是所有表都有了？
        if (
            self.m_totalchart_filePicker1.GetPath()
            and self.m_noneed_chart_filePicker111.GetPath()
        ):
            self.go_for_validate()
        # event.Skip()

    def on_nocost_chart_file_change(self, event):
        self.reset_read_result()
        self.m_noneed_picker_staticText211.SetLabel(
            Path(self.m_noneed_chart_filePicker111.GetPath()).name
        )
        if (
            self.m_totalchart_filePicker1.GetPath()
            and self.m_noneed_chart_filePicker111.GetPath()
        ):
            self.go_for_validate()
        event.Skip()

    def on_read_data_school_change(self, event):
        now_school = self.m_school_choice1.GetStringSelection().replace(
            "（已读入）", ""
        )
        # prt(now_school)
        # prt(self.sp_school_stats[now_school])
        if stat := self.sp_school_stats.get(now_school, {}):
            self.short_result_text.SetLabelText(
                "【检查结果】总人数：{total:3d}\t\t需缴费学生人数：{shouldpay:3d}\t免缴费学生人数：{exempt:3d}\t错误数据：{error:3d}\t错误占比：{error_rate:.2f}%".format(
                    **self.sp_school_stats[now_school]
                )
            )
        else:
            self.reset_read_result()
        event.Skip()

    def on_read_data_save_button_click(self, event):
        now_school = self.m_school_choice1.GetStringSelection().replace(
            "（已读入）", ""
        )

        self.sp_school_pscmanagers[now_school] = self.current_pscmanager
        self.sp_school_stats[now_school] = self.current_psc_stats
        self.sp_school_errors[now_school] = self.current_psc_errors

        if self.store_path is None:
            self.initial_store_path()

        (now_school_store_dir := (self.store_path / now_school).absolute()).mkdir(  # type: ignore
            parents=True, exist_ok=True
        )
        self.current_pscmanager.save_to_json(now_school_store_dir / "students.json")
        with open(now_school_store_dir / "errors.json", "w") as f:
            json.dump(self.current_psc_errors, f)

        self.m_school_choice1.SetString(
            self.m_school_choice1.GetSelection(), now_school + "（已读入）"
        )
        self.reset_page1_all()

        if self.m_school_choice1.GetSelection() == self.m_school_choice1.GetCount() - 1:
            for i in range(self.m_school_choice1.GetCount()):
                if "已读入" not in self.m_school_choice1.GetString(i):
                    self.m_school_choice1.SetSelection(i)
                    break
            wx.MessageDialog(self, "所有学院的数据皆已导入", "提示").ShowModal()
        else:
            self.m_school_choice1.SetSelection(self.m_school_choice1.GetSelection() + 1)

        event.Skip()

    def on_page1_notebook_changed(self, event):
        total_count = {"WC": {"T": 0, "G": 0}, "DX": {"T": 0, "G": 0}}
        self.m_export_total_dataViewListCtrl2.DeleteAllItems()
        if self.m_notebook1.GetSelection() == 2:
            for school, stat in self.sp_school_stats.items():
                # {"total": 0, "shouldpay": 0, "exempt": 0, "error": 0, "grad": 0, "error_rate": 0.0"}
                if stat:
                    # self.m_export_total_dataViewListCtrl2
                    self.m_export_total_dataViewListCtrl2.AppendItem(
                        [
                            school,
                            stat["total"],
                            stat["shouldpay"],
                            stat["exempt"],
                            stat["error"],
                            str(25 * int(stat["shouldpay"])),
                        ]
                    )
                    if school in SPERATE_SCHOOL_ZONE["坞城"]:
                        total_count["WC"]["T"] += stat["total"]
                        total_count["WC"]["G"] += stat["grad"]
                    elif school in SPERATE_SCHOOL_ZONE["东山"]:
                        total_count["DX"]["T"] += stat["total"]
                        total_count["DX"]["G"] += stat["grad"]
                    else:
                        # 怎么可能呢
                        pass

                    self.m_export_total_dataViewListCtrl2.Update()
                    # self.m_checkin_dataViewListColumn7.
            self.m_totalexport_staticText14.SetLabelText(
                "总报名人数：{:3d}人\t\t坞城校区：{:3d}人\t\t东山校区：{:3d}人".format(
                    total_count["WC"]["T"] + total_count["DX"]["T"],
                    total_count["WC"]["T"],
                    total_count["DX"]["T"],
                )
            )
            self.m_wucheng_stats_staticText152.SetLabelText(
                "\t其中，坞城校区——\t非毕业年级：{:3d}人\t\t毕业年级：{:3d}人".format(
                    total_count["WC"]["T"] - total_count["WC"]["G"],
                    total_count["WC"]["G"],
                )
            )
            self.m_export_total_dataViewListCtrl2.Update()
            self.m_totalexport_staticText14.Update()
            self.m_wucheng_stats_staticText152.Update()

        event.Skip()

    def on_page2_notebook_changed(self, event):
        event.Skip()

    def on_wrong_data_item_editing_done(self, event):
        event.Skip()

    def on_wrong_data_item_editing_start(self, event):
        event.Skip()

    def on_wrong_data_item_value_changed(self, event):
        event.Skip()

    def on_grid_cell_change(self, event):
        event.Skip()

    def on_grid_cell_left_click(self, event):
        event.Skip()

    def on_grid_cell_right_click(self, event):
        event.Skip()

    def on_export_each_school_button_click(self, event):
        event.Skip()

    def on_export_total_chart_button_click(self, event):
        event.Skip()

    # Virtual image path resolution method. Override this in your derived class.
    def img_wrapper(self, bitmap_path):
        return bitmap_path


if __name__ == "__main__":
    app = wx.App(False)
    frame = CostCheckerMainFrame(None)
    frame.Show(True)
    app.MainLoop()
