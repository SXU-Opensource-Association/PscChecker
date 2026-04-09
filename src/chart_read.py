# -*- coding: utf-8 -*-

import pandas as pd
import json
from typing import List, Dict, Optional, Any, Set
from pathlib import Path


class CheckPerson:
    """存储单个学生普通话测试信息的类"""

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
        should_pay_cost: str = "25",
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
        self.should_pay_cost = str(
            should_pay_cost if should_pay_cost in ["25", "35", "0"] else ""
        )

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
            "应缴费款": self.should_pay_cost,
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
            should_pay_cost=data.get("应缴费款", ""),
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

    def load_from_excel(self, file_path: str) -> None:
        """
        从Excel文件加载学生数据

        Args:
            file_path: Excel文件路径
        """
        try:
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
            self.students = []
            for _, row in df.iterrows():
                row = row.to_dict()
                student_data = {}
                for col_name, attr_name in self.TOTAL_CHART_COLUMN_MAPPING.items():
                    if col_name in df.columns:
                        student_data[attr_name] = row[col_name]

                self.students.append(CheckPerson(**student_data))

        except Exception as e:
            raise ValueError(f"加载Excel文件失败: {str(e)}")

    def save_to_excel(self, file_path: str) -> None:
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

    def load_from_json(self, file_path: str) -> None:
        """
        从JSON文件加载学生数据

        Args:
            file_path: JSON文件路径
        """
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.students = [CheckPerson.from_dict(item) for item in data]

        except Exception as e:
            raise ValueError(f"加载JSON文件失败: {str(e)}")

    def save_to_json(self, file_path: str) -> None:
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
    ) -> Dict[str, List[Dict[str, str]]]:
        """
        校验总表与缴费名单的一致性

        Args:
            exempt_file: 免缴费名单Excel文件路径
            fee_file: 缴费名单Excel文件路径

        Returns:
            dict: 包含各种错误类型的字典
        """
        # 1. 读取免缴费名单
        exempt_df: pd.DataFrame = pd.read_excel(exempt_file, dtype=str).fillna("")
        exempt_df = exempt_df[["序号", "学院", "姓名", "学号", "备注"]]

        # 2. 读取缴费名单
        fee_df: pd.DataFrame = pd.read_excel(fee_file, dtype=str).fillna("")
        fee_df = fee_df[["序号", "姓名", "性别", "院系", "学号", "缴费金额"]]

        # 3. 构建索引：学号 -> 姓名映射
        exempt_index: Dict[str, Dict[str, str]] = {}
        for _, row in exempt_df.iterrows():
            student_id = str(row["学号"]).strip()
            if student_id:
                exempt_index[student_id] = {
                    "姓名": str(row["姓名"]).strip(),
                    "学院": str(row["学院"]).strip(),
                    "备注": str(row["备注"]).strip(),
                }

        fee_index: Dict[str, Dict[str, str]] = {}
        for _, row in fee_df.iterrows():
            student_id = str(row["学号"]).strip()
            if student_id:
                fee_index[student_id] = {
                    "姓名": str(row["姓名"]).strip(),
                    "性别": str(row["性别"]).strip(),
                    "院系": str(row["院系"]).strip(),
                    "缴费金额": str(row["缴费金额"]).strip(),
                }

        # 4. 准备结果字典
        errors = {
            "missing_in_both": [],  # 在总表中但不在任何名单中
            "fee_amount_error": [],  # 缴费金额有误
            "in_fee_not_in_total": [],  # 在缴费名单但不在总表中
            "in_exempt_not_in_total": [],  # 在免缴费名单但不在总表中
            "name_mismatch": [],  # 姓名不一致
            "student_id_mismatch": [],  # 学号不一致（用于检测重复学号）
        }

        # 5. 遍历总表中的每个人员
        total_name_map: Dict[str, str] = {}

        for person in self.students:
            person.student_id = person.student_id.strip()
            person.name = person.name.strip()
            person.id_number = person.id_number.strip()

            total_name_map[person.student_id] = person.name

        # 6. 检查总表中的人是否在缴费名单中
        for person in self.students:
            # student_id = person.student_id.strip()
            # name = person.name.strip()

            # # 如果学号为空，跳过
            # if not student_id:
            #     continue

            in_exempt = person.student_id in exempt_index
            in_fee = person.student_id in fee_index

            # 情况1: 在总表中但不在任何名单中
            if not in_exempt and not in_fee:
                errors["missing_in_both"].append(
                    {
                        "学号": person.student_id,
                        "姓名": person.name,
                        "身份证": person.id_number,
                        "原因": "未在免缴费名单或缴费名单中找到",
                    }
                )

            # 情况2: 在免缴费名单中
            if in_exempt:
                exempt_info = exempt_index[person.student_id]
                # 检查姓名是否一致
                if exempt_info["姓名"] != person.name:
                    errors["name_mismatch"].append(
                        {
                            "学号": person.student_id,
                            "姓名": person.name,
                            "身份证": person.id_number,
                            "原因": "免缴费名单与总表姓名不一致",
                        }
                    )

                person.should_pay_cost = "0"
                # # 检查应缴费金额是否正确
                # if person.should_pay_cost != "0":
                #     errors["fee_amount_error"].append(
                #         {
                #             "学号": person.student_id,
                #             "姓名": person.name,
                #             "身份证": person.id_number,
                #             "原因": "免缴费人员应缴费金额应为0",
                #         }
                #     )

            # 情况3: 在缴费名单中
            if in_fee:
                fee_info = fee_index[person.student_id]
                # 检查姓名是否一致
                if fee_info["姓名"] != person.name:
                    errors["name_mismatch"].append(
                        {
                            "学号": person.student_id,
                            "总表姓名": person.name,
                            "缴费名单姓名": fee_info["姓名"],
                            "原因": "姓名不一致",
                        }
                    )

                # 检查缴费金额是否正确
                fee_amount = fee_info["缴费金额"]

                if fee_amount not in ["25", "35"]:
                    errors["fee_amount_error"].append(
                        {
                            "学号": person.student_id,
                            "姓名": person.name,
                            "缴费名单金额": fee_amount,
                            "原因": "缴费金额应为25或35元",
                        }
                    )

        # 7. 检查缴费名单中但不在总表中的人
        for student_id, fee_info in fee_index.items():
            if student_id not in total_name_map:
                errors["in_fee_not_in_total"].append(
                    {
                        "学号": student_id,
                        "姓名": fee_info["姓名"],
                        "性别": fee_info.get("性别", ""),
                        "院系": fee_info.get("院系", ""),
                        "缴费金额": fee_info["缴费金额"],
                        "原因": "在缴费名单中但未在总表中找到",
                    }
                )

        # 8. 检查免缴费名单中但不在总表中的人
        for student_id, exempt_info in exempt_index.items():
            if student_id not in total_name_map:
                errors["in_exempt_not_in_total"].append(
                    {
                        "学号": student_id,
                        "姓名": exempt_info["姓名"],
                        "学院": exempt_info.get("学院", ""),
                        "备注": exempt_info.get("备注", ""),
                        "原因": "在免缴费名单中但未在总表中找到",
                    }
                )

        # 9. 清理空列表
        for key in list(errors.keys()):
            if not errors[key]:
                del errors[key]

        return errors

    def auto_fill_should_pay_cost(self):
        """根据人员类型自动填充应缴费金额"""
        for person in self.students:
            # 默认为学生
            cost = "25"

            # 判断是否为教师
            is_teacher = False
            if (
                (
                    person.occupation
                    and any(
                        kw in person.occupation for kw in ["教师", "教职工", "教工"]
                    )
                )
                or (
                    person.department
                    and any(kw in person.department for kw in ["教师", "教工"])
                )
                or (
                    person.class_name
                    and any(kw in person.class_name for kw in ["教师", "教工"])
                )
            ):
                is_teacher = True

            person.should_pay_cost = "35" if is_teacher else "25"

    def get_summary_stats(self) -> Dict[str, int]:
        """获取统计摘要"""
        stats = {"total": len(self.students), "students": 0, "teachers": 0, "exempt": 0}

        for person in self.students:
            if person.should_pay_cost == "25":
                stats["students"] += 1
            elif person.should_pay_cost == "35":
                stats["teachers"] += 1
            elif person.should_pay_cost == "0":
                stats["exempt"] += 1

        return stats


# 使用示例
if __name__ == "__main__":
    # 创建管理器实例
    manager = PscPersonManager()

    # 示例：从Excel读取数据
    # manager.load_from_excel("mandarin_test_list.xlsx")

    # 示例：添加学生
    student1 = CheckPerson(
        name="张三",
        gender="男",
        ethnicity="汉族",
        id_type="身份证",
        id_number="123456789012345678",
        occupation="学生",
        organization="XX大学",
        phone="13800138000",
        student_id="20230001",
        class_name="计算机科学与技术2023级1班",
        department="计算机学院",
        is_graduating=False,
        contact_address="XX市XX区XX路XX号",
        mailing_address="XX市XX区XX路XX号",
        postal_code="100000",
        birth_province="北京市",
        birth_city="北京市",
        birth_district="海淀区",
        current_province="北京市",
        current_city="北京市",
        current_district="海淀区",
    )

    manager.add_student(student1)

    # 保存为JSON（高效的数据存储格式）
    manager.save_to_json("mandarin_test_data.json")

    # 从JSON加载
    manager_new = PscPersonManager()
    manager_new.load_from_json("mandarin_test_data.json")

    # 导出为Excel
    manager_new.save_to_excel("mandarin_test_output.xlsx")

    print(f"成功处理 {len(manager_new.get_students())} 名学生数据")
