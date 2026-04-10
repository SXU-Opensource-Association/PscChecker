# -*- coding: utf-8 -*-

import pandas as pd
import json
from typing import List, Dict, Optional, Any, Set, Iterable, Mapping
from pathlib import Path

from datetime import datetime, timedelta, date
from dataclasses import dataclass
from rich import print as prt


# 校验码验证
id_validation_weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
id_validation_check_codes = ["1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2"]


def validate_id_number(id_card_n: str):
    """
    初步验证18位中国身份证号码是否格式正确。
    """
    if not isinstance(id_card_n, str):
        return False

    id_card_n = id_card_n.strip().upper()

    if len(id_card_n) != 18:
        return False

    # 检查前17位是否全为数字
    if not id_card_n[:17].isdigit():
        return False

    # 检查最后一位是否为数字或大写X
    if not (id_card_n[17].isdigit() or id_card_n[17] == "X"):
        return False

    # 提取出生年月日
    try:
        year = int(id_card_n[6:10])
        month = int(id_card_n[10:12])
        day = int(id_card_n[12:14])

        birthday = date(year, month, day)
    except (ValueError, OverflowError):
        return False

    # 简单范围检查
    if not (date(1949, 10, 1) < birthday < datetime.now().date()):
        return False

    return (
        id_card_n[17]
        == id_validation_check_codes[
            sum(id_validation_weights[i] * int(id_card_n[i]) for i in range(17)) % 11
        ]
    )


@dataclass
class PersonStatus:
    OK = 0
    FINE_WITH_ID_ERROR = 1
    FINE_WITH_NAME_ERROR = 2
    ERROR = 3

    name: str
    id: str

    def __init__(self, name: str, id: Any):
        self.id = str(id).strip()
        self.name = name.strip()

    def compare(self, other: "PersonStatus") -> int:
        if self.id == other.id:
            if self.name == other.name:
                return PersonStatus.OK
            else:
                return PersonStatus.FINE_WITH_NAME_ERROR
        else:
            if self.name == other.name:
                return PersonStatus.FINE_WITH_ID_ERROR
            else:
                return PersonStatus.ERROR

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, PersonStatus):
            return False
        return self.compare(other) == PersonStatus.OK

    # def get_status_text(self) -> str:
    #     if self.status == PersonStatus.OK:
    #         return _("存在")
    #     elif self.status == PersonStatus.FINE_WITH_ID_ERROR:
    #         return _("姓名存在*学号错误")
    #     elif self.status == PersonStatus.FINE_WITH_NAME_ERROR:
    #         return _("学号存在*姓名错误")
    #     else:
    #         return _("查无此人")

    def get_name(self) -> str:
        return self.name

    def get_id(self) -> str:
        return self.id

    def __repr__(self) -> str:
        return f"PersonStatus(name={self.name}, id={self.id})"

    def __str__(self) -> str:
        return f"{self.name} {self.id}"


class CheckPerson:
    """存储单个学生普通话测试信息的类"""

    name: str
    """考生姓名"""
    gender: str
    """考生性别"""
    ethnicity: str
    """考生民族"""
    id_type: str
    """证件类型"""
    id_number: str
    """证件编号"""
    occupation: str
    """从事职业"""
    organization: str
    """所在单位"""
    phone: str
    """联系电话"""
    student_id: str
    """考生学号"""
    class_name: str
    """考生班级"""
    department: str
    """考生院系"""
    is_graduating: bool
    """是否为毕业年级"""
    contact_address: str
    """联系地址"""
    mailing_address: str
    """邮寄地址"""
    postal_code: str
    """邮政编码"""
    birth_province: str
    """出生所在省"""
    birth_city: str
    """出生所在城市"""
    birth_district: str
    """出生所在县(区)"""
    current_province: str
    """现居住省"""
    current_city: str
    """现居住城市"""
    current_district: str
    """现居住县(区)"""
    should_pay_cost: Optional[bool]
    """是否应缴费"""

    def __init__(
        self,
        name: str = "",
        gender: str = "",
        ethnicity: str = "",
        id_type: str = "",
        id_number: str = "",
        occupation: str = "",
        organization: str = "",
        phone: str = "",
        student_id: str = "",
        class_name: str = "",
        department: str = "",
        is_graduating: bool = False,
        contact_address: str = "",
        mailing_address: str = "",
        postal_code: str = "",
        birth_province: str = "",
        birth_city: str = "",
        birth_district: str = "",
        current_province: str = "",
        current_city: str = "",
        current_district: str = "",
        should_pay_cost: Optional[bool] = None,
    ):
        """
        初始化学生信息

        Args:
            name: 考生姓名
            gender: 考生性别
            ethnicity: 考生民族
            id_type: 证件类型
            id_number: 证件编号
            occupation: 从事职业
            organization: 所在单位
            phone: 联系电话
            student_id: 考生学号
            class_name: 考生班级
            department: 考生院系
            is_graduating: 是否为毕业年级
            contact_address: 联系地址
            mailing_address: 邮寄地址
            postal_code: 邮政编码
            birth_province: 出生所在省
            birth_city: 出生所在城市
            birth_district: 出生所在县(区)
            current_province: 现居住省
            current_city: 现居住城市
            current_district: 现居住县(区)
        """
        self.name = name
        self.gender = gender
        self.ethnicity = ethnicity
        self.id_type = id_type
        self.id_number = str(id_number)
        self.occupation = occupation
        self.organization = organization
        self.phone = str(phone)
        self.student_id = str(student_id)
        self.class_name = class_name
        self.department = department
        self.is_graduating = is_graduating
        self.contact_address = contact_address
        self.mailing_address = mailing_address
        self.postal_code = str(postal_code)
        self.birth_province = birth_province
        self.birth_city = birth_city
        self.birth_district = birth_district
        self.current_province = current_province
        self.current_city = current_city
        self.current_district = current_district
        self.should_pay_cost = should_pay_cost

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "考生姓名": self.name,
            "考生性别": self.gender,
            "考生民族": self.ethnicity,
            "证件类型": self.id_type,
            "证件编号": self.id_number,
            "从事职业": self.occupation,
            "所在单位": self.organization,
            "联系电话": self.phone,
            "考生学号": self.student_id,
            "考生班级": self.class_name,
            "考生院系": self.department,
            "是否为毕业年级": self.is_graduating,
            "联系地址": self.contact_address,
            "邮寄地址": self.mailing_address,
            "邮政编码": self.postal_code,
            "出生所在省": self.birth_province,
            "出生所在城市": self.birth_city,
            "出生所在县(区)": self.birth_district,
            "现居住省": self.current_province,
            "现居住城市": self.current_city,
            "现居住县(区)": self.current_district,
            "是否应缴费": self.should_pay_cost,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CheckPerson":
        """从字典创建实例"""
        return cls(
            name=data.get("考生姓名", ""),
            gender=data.get("考生性别", ""),
            ethnicity=data.get("考生民族", ""),
            id_type=data.get("证件类型", ""),
            id_number=data.get("证件编号", ""),
            occupation=data.get("从事职业", ""),
            organization=data.get("所在单位", ""),
            phone=data.get("联系电话", ""),
            student_id=data.get("考生学号", ""),
            class_name=data.get("考生班级", ""),
            department=data.get("考生院系", ""),
            is_graduating=bool(
                True if ("是" in (i := data.get("是否为毕业年级", False))) else i
            ),
            contact_address=data.get("联系地址", ""),
            mailing_address=data.get("邮寄地址", ""),
            postal_code=data.get("邮政编码", ""),
            birth_province=data.get("出生所在省", ""),
            birth_city=data.get("出生所在城市", ""),
            birth_district=data.get("出生所在县(区)", ""),
            current_province=data.get("现居住省", ""),
            current_city=data.get("现居住城市", ""),
            current_district=data.get("现居住县(区)", ""),
            should_pay_cost=data.get("是否应缴费", ""),
        )

    def __repr__(self) -> str:
        return f"CheckPerson(name='{self.name}', student_id='{self.student_id}')"


class PscPersonManager:
    """普通话测试名单管理器"""

    # Excel表格的列名映射
    TOTAL_CHART_COLUMN_MAPPING = {
        "考生姓名": "name",
        "考生性别": "gender",
        "考生民族": "ethnicity",
        "证件类型": "id_type",
        "证件编号": "id_number",
        "从事职业": "occupation",
        "所在单位": "organization",
        "联系电话": "phone",
        "考生学号": "student_id",
        "考生班级": "class_name",
        "考生院系": "department",
        "是否为毕业年级": "is_graduating",
        "联系地址": "contact_address",
        "邮寄地址": "mailing_address",
        "邮政编码": "postal_code",
        "出生所在省": "birth_province",
        "出生所在城市": "birth_city",
        "出生所在县(区)": "birth_district",
        "现居住省": "current_province",
        "现居住城市": "current_city",
        "现居住县(区)": "current_district",
    }

    def __init__(self):
        self.students: List[CheckPerson] = []

    @classmethod
    def load_from_excel(cls, file_path: str | Path) -> "PscPersonManager":
        """
        从Excel文件加载学生数据

        Args:
            file_path: Excel文件路径
        """
        try:
            instance = cls()
            # 读取Excel文件
            df = pd.read_excel(file_path, dtype=str)

            # 处理布尔值列
            if "是否为毕业年级" in df.columns:
                df["是否为毕业年级"] = df["是否为毕业年级"].map(
                    lambda x: (True if ("是" in str(x).lower()) else False)
                )

            # 清理数据：将NaN转换为空字符串
            df: pd.DataFrame = df.fillna("")

            # 转换为CheckPerson对象列表
            instance.students = []
            for _, row in df.iterrows():
                row = row.to_dict()
                student_data = {}
                for col_name, attr_name in instance.TOTAL_CHART_COLUMN_MAPPING.items():
                    if col_name in df.columns:
                        student_data[attr_name] = row[col_name]

                instance.students.append(CheckPerson(**student_data))
            return instance
        except Exception as e:
            raise ValueError(f"加载Excel文件失败: {str(e)}")

    def save_to_excel(self, file_path: str | Path) -> None:
        """
        保存学生数据到Excel文件

        Args:
            file_path: 输出Excel文件路径
        """
        if not self.students:
            # 创建空的DataFrame
            df = pd.DataFrame(columns=list(self.TOTAL_CHART_COLUMN_MAPPING.keys()))
        else:
            # 转换为DataFrame
            data_list = [student.to_dict() for student in self.students]
            df = pd.DataFrame(
                data_list, columns=list(self.TOTAL_CHART_COLUMN_MAPPING.keys())
            )

        # 创建ExcelWriter并设置列格式为文本
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

            # 获取工作表对象
            worksheet = writer.sheets["Sheet1"]

            # 设置所有列为文本格式（防止Excel自动转换数字格式）
            text_format = writer.book.add_format({"num_format": "@"})
            for col_num, col_name in enumerate(df.columns):
                # 应用文本格式到整列
                worksheet.set_column(col_num, col_num, None, text_format)

    @classmethod
    def load_from_json(cls, file_path: str | Path) -> "PscPersonManager":
        """
        从JSON文件加载学生数据

        Args:
            file_path: JSON文件路径
        """
        try:
            instance = cls()
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            instance.students = [CheckPerson.from_dict(item) for item in data]
            return instance

        except Exception as e:
            raise ValueError(f"加载JSON文件失败: {str(e)}")

    def save_to_json(self, file_path: str | Path) -> None:
        """
        保存学生数据到JSON文件

        Args:
            file_path: 输出JSON文件路径
        """
        data = [student.to_dict() for student in self.students]

        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def add_student(self, student: CheckPerson) -> None:
        """添加学生"""
        self.students.append(student)

    def get_students(self) -> List[CheckPerson]:
        """获取所有学生"""
        return self.students.copy()

    def clear(self) -> None:
        """清空所有学生数据"""
        self.students.clear()

    def validate_against_fee_lists(
        self, exempt_file: str, fee_file: str
    ) -> Iterable[Dict[str, Any]]:
        """
        校验总表与缴费名单的一致性

        Args:
            exempt_file: 免缴费名单Excel文件路径
            fee_file: 缴费名单Excel文件路径

        Returns:
            dict: 包含各种错误类型的字典
        """
        # 1. 读取免缴费名单
        exempt_df: pd.DataFrame = pd.read_excel(
            exempt_file, skiprows=2, dtype=str
        ).fillna("")
        exempt_df = exempt_df[2:][["序号", "学院", "姓名", "学号", "备注"]]

        # 2. 读取缴费名单
        fee_df: pd.DataFrame = pd.read_excel(fee_file, skiprows=1, dtype=str).fillna("")
        fee_df = fee_df[1:][["序号", "姓名", "性别", "院系", "学号", "缴费金额"]]

        # 3. 构建索引
        exempt_ids: List[str] = []
        exempt_names: List[str] = []
        for _, row in exempt_df.iterrows():
            exempt_ids.append(str(row["学号"]).strip())
            exempt_names.append(str(row["姓名"]).strip())

        fee_ids: List[str] = []
        fee_names: List[str] = []
        fee_costs: List[str] = []
        for _, row in fee_df.iterrows():
            fee_ids.append(str(row["学号"]).strip())
            fee_names.append(str(row["姓名"]).strip())
            fee_costs.append(str(row["缴费金额"]).strip())

        # 4. 准备结果字典
        # errors = {
        #     "missing_in_both": [],  # 在总表中但不在任何名单中
        #     "fee_amount_error": [],  # 缴费金额有误
        #     "in_fee_not_in_total": [],  # 在缴费名单但不在总表中
        #     "in_exempt_not_in_total": [],  # 在免缴费名单但不在总表中
        #     "name_mismatch": [],  # 姓名不一致
        #     "student_id_mismatch": [],  # 学号不一致（用于检测重复学号）
        # }

        # 5. 遍历总表中的每个人员
        total_ids: List[str] = []
        total_names: List[str] = []

        for i, person in enumerate(self.students):
            person.student_id = person.student_id.strip()
            person.name = person.name.strip()
            person.id_number = person.id_number.strip()
            if (not validate_id_number(person.id_number)) and (
                person.id_type == "身份证"
            ):
                yield {
                    "总表项目序号": i,
                    "学号": person.student_id,
                    "姓名": person.name,
                    "身份证": person.id_number,
                    "需缴费": None,
                    "原因": "身份证号格式错误",
                }

            total_ids.append(person.student_id)
            total_names.append(person.name)

        exempt_checked_index = []
        fee_checked_index = []

        # 6. 检查总表中的人是否在缴费名单中
        for i in range(len(self.students)):
            # student_id = person.student_id.strip()
            # name = person.name.strip()

            # # 如果学号为空，跳过
            # if not student_id:
            #     continue

            person = self.students[i]

            in_exempt = (person.student_id in exempt_ids) + (
                person.name in exempt_names
            )
            in_fee = (person.student_id in fee_ids) + (person.name in fee_names)

            # 情况1: 在总表中但不在任何名单中
            if not in_exempt and not in_fee:
                yield {
                    "总表项目序号": i,
                    "学号": person.student_id,
                    "姓名": person.name,
                    "身份证": person.id_number,
                    "需缴费": None,
                    "原因": "不在缴费或免缴名单中",
                }

            # 情况2: 在免缴费名单中
            elif in_exempt:
                if in_exempt == 1:
                    if person.student_id in exempt_ids:
                        exempt_checked_index.append(
                            _eii := exempt_ids.index(person.student_id)
                        )
                        yield {
                            "总表项目序号": i,
                            "学号": person.student_id,
                            "姓名": person.name,
                            "身份证": person.id_number,
                            "需缴费": False,
                            "原因": "免缴费名单中的【姓名：{}】与总表不一致".format(
                                exempt_names[_eii]
                            ),
                        }
                    else:
                        exempt_checked_index.append(
                            _eni := exempt_names.index(person.name)
                        )
                        yield {
                            "总表项目序号": i,
                            "学号": person.student_id,
                            "姓名": person.name,
                            "身份证": person.id_number,
                            "需缴费": False,
                            "原因": "免缴费名单中的【学号：{}】与总表不一致".format(
                                exempt_ids[_eni]
                            ),
                        }
                else:
                    _exe_nme = exempt_names[_eii := exempt_ids.index(person.student_id)]
                    exempt_checked_index.append(_eii)
                    if _exe_nme == person.name:
                        person.should_pay_cost = False
                    else:
                        yield {
                            "总表项目序号": i,
                            "学号": person.student_id,
                            "姓名": person.name,
                            "身份证": person.id_number,
                            "需缴费": False,
                            "原因": "免缴费名单中的【姓名{}】与总表不一致".format(
                                _exe_nme
                            ),
                        }
            else:
                person.should_pay_cost = True
                if in_fee == 1:
                    if person.student_id in fee_ids:
                        fee_checked_index.append(fee_ids.index(person.student_id))
                    else:
                        fee_checked_index.append(fee_names.index(person.name))
                else:
                    fee_checked_index.append(fee_ids.index(person.student_id))

            # # 情况3: 在缴费名单中
            # if in_fee:
            #     if in_fee == 1:
            #         if person.student_id in fee_ids:
            #             fee_checked_index.append(
            #                 _fii := fee_ids.index(person.student_id)
            #             )
            #             yield {
            #                 "总表项目序号": i,
            #                 "学号": person.student_id,
            #                 "姓名": person.name,
            #                 "身份证": person.id_number,
            #                 "需缴费": True,
            #                 "原因": "缴费名单中的【姓名：{}】与总表不一致".format(
            #                     fee_names[_fii]
            #                 ),
            #             }
            #         else:
            #             yield {
            #                 "总表项目序号": i,
            #                 "学号": person.student_id,
            #                 "姓名": person.name,
            #                 "身份证": person.id_number,
            #                 "需缴费": True,
            #                 "原因": "缴费名单中的【学号：{}】与总表不一致".format(
            #                     fee_ids[fee_names.index(person.name)]
            #                 ),
            #             }
            #     else:
            #         _fee_nme = fee_names[fee_ids.index(person.student_id)]
            #         _fee_id = fee_ids[fee_names.index(person.name)]
            #         if _fee_nme == person.name:
            #             if _fee_id == person.student_id:
            #                 person.should_pay_cost = True
            #             else:
            #                 yield {
            #                     "总表项目序号": i,
            #                     "学号": person.student_id,
            #                     "姓名": person.name,
            #                     "身份证": person.id_number,
            #                     "需缴费": True,
            #                     "原因": "缴费名单中的【学号{}】与总表不一致".format(
            #                         _fee_id
            #                     ),
            #                 }
            #         else:
            #             if _fee_id == person.student_id:
            #                 yield {
            #                     "总表项目序号": i,
            #                     "学号": person.student_id,
            #                     "姓名": person.name,
            #                     "身份证": person.id_number,
            #                     "需缴费": True,
            #                     "原因": "缴费名单中的【姓名{}】与总表不一致".format(
            #                         _fee_nme
            #                     ),
            #                 }
            #             else:
            #                 yield {
            #                     "总表项目序号": i,
            #                     "学号": person.student_id,
            #                     "姓名": person.name,
            #                     "身份证": person.id_number,
            #                     "需缴费": True,
            #                     "原因": "缴费名单中的【学号{}】【姓名{}】皆与总表不一致".format(
            #                         _fee_id, _fee_nme
            #                     ),
            #                 }

        # 7. 检查缴费名单中但不在总表中的人
        for i in range(len(fee_ids)):
            if i not in fee_checked_index:
                self.students.append(
                    CheckPerson(
                        name=fee_names[i],
                        student_id=fee_ids[i],
                        should_pay_cost=True,
                    )
                )
                yield {
                    "总表项目序号": len(self.students) - 1,
                    "学号": fee_ids[i],
                    "姓名": fee_names[i],
                    "身份证": "??",
                    "需缴费": True,
                    "原因": "在缴费名单中但未在总表中找到",
                }

        # 8. 检查免缴费名单中但不在总表中的人
        for i in range(len(exempt_ids)):
            if i not in exempt_checked_index:
                self.students.append(
                    CheckPerson(
                        name=exempt_names[i],
                        student_id=exempt_ids[i],
                        should_pay_cost=False,
                    )
                )
                yield {
                    "总表项目序号": len(self.students) - 1,
                    "学号": exempt_ids[i],
                    "姓名": exempt_names[i],
                    "身份证": "??",
                    "需缴费": False,
                    "原因": "在免缴费名单中但未在总表中找到",
                }

    def get_summary_stats(self):
        """获取统计摘要"""
        stats = {"total": len(self.students), "shouldpay": 0, "exempt": 0, "error": 0, "grad": 0, "error_rate": 0.0}

        for person in self.students:
            if person.should_pay_cost is None:
                stats["error"] += 1
            else:
                stats["grad"] += person.is_graduating
                if person.should_pay_cost:
                    stats["shouldpay"] += 1
                else:
                    stats["exempt"] += 1

        stats["error_rate"] = stats["error"] * 100 / stats["total"]

        return stats
