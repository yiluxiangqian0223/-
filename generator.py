#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
考场图像单生成器
用于生成高校考试的图像信息核对单
"""

import os
import sys
import logging
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import pandas as pd
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def register_chinese_fonts() -> str:
    """
    注册中文字体
    
    Returns:
        str: 注册成功的字体名称，如果失败则返回默认字体
    """
    # 尝试多个常见中文字体路径
    font_configs = [
        ('/System/Library/Fonts/STHeiti Light.ttc', 'STHeiti-Light'), #Mac
        ('/System/Library/Fonts/STHeiti Medium.ttc', 'STHeiti-Medium'),
        ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
        ('C:/Windows/Fonts/simsun.ttc', 'SimSun'), #Windows
        ('C:/Windows/Fonts/msyh.ttc', 'MicrosoftYaHei'),
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 'DejaVuSans'),
    ]
    
    for font_path, font_name in font_configs:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                logger.info(f"✅ 中文字体注册成功：{font_path}")
                return font_name
            except Exception as e:
                logger.warning(f"❌ 字体注册失败 {font_path}: {e}")
                continue
    
    logger.warning("⚠️ 未能找到合适的中文字体，使用默认字体")
    return 'Helvetica'


def create_vertical_text(text: str) -> str:
    """
    创建竖排文字（用换行符实现）
    
    Args:
        text: 输入文字
        
    Returns:
        str: 竖排格式的文字
    """
    return '<br/>'.join(list(text))


class ExamImageGenerator:
    """考场图像单生成器"""
    
    def __init__(
        self, 
        data_dir: str, 
        photo_dir: str, 
        output_dir: str = './output_exam_images'
    ):
        """
        初始化生成器
        
        Args:
            data_dir: 数据目录路径
            photo_dir: 照片目录路径  
            output_dir: 输出目录路径
        """
        self.data_dir = Path(data_dir)
        self.photo_dir = Path(photo_dir)
        self.output_dir = Path(output_dir)
        
        # 创建输出目录
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # 注册字体
        self.chinese_font = register_chinese_fonts()
        
        # 加载考生数据
        self.df = None
        self.photo_map = {}
        self.exam_rooms = {}
        self.load_data()
        
        # 页面配置
        self.students_per_page = 30
        self.rows = 5
        self.cols = 6
        
        # 照片尺寸
        self.photo_width = 2.16 * cm
        self.photo_height = 2.74 * cm
        
        # 左侧竖排区域宽度
        self.side_width = 0.8 * cm
        
        # 样式
        self.styles = getSampleStyleSheet()
        self._setup_styles()
        
        logger.info(f"✅ 生成器初始化完成，输出目录: {self.output_dir}")
    
    def load_data(self) -> None:
        """加载考生数据和照片信息"""
        excel_path = self.data_dir / '2026kctx.xlsx'
        
        if not excel_path.exists():
            raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
        
        self.df = pd.read_excel(excel_path)
        logger.info(f"✅ 已加载 {len(self.df)} 名考生信息")
        
        # 检查照片
        if not self.photo_dir.exists():
            raise FileNotFoundError(f"照片目录不存在: {self.photo_dir}")
        
        photo_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff'}
        self.photo_map = {}
        
        for photo_file in self.photo_dir.iterdir():
            if photo_file.suffix.lower() in photo_extensions:
                id_card = photo_file.stem  # 移除扩展名得到身份证号
                self.photo_map[id_card] = str(photo_file)
        
        logger.info(f"✅ 已加载 {len(self.photo_map)} 张照片")
        
        # 按考场分组
        self.exam_rooms = {}
        for room_num in self.df['考场号'].unique():
            room_data = self.df[self.df['考场号'] == room_num].sort_values('座位号')
            self.exam_rooms[room_num] = room_data
        
        logger.info(f"✅ 共 {len(self.exam_rooms)} 个考场")
    
    def _setup_styles(self) -> None:
        """设置样式"""
        # 大标题样式
        self.title_style = ParagraphStyle(
            'Title',
            parent=self.styles['Normal'],
            fontSize=17,
            alignment=TA_CENTER,
            fontName=self.chinese_font,
            spaceAfter=15,
            leading=24
        )
        
        # 信息行样式
        self.info_style = ParagraphStyle(
            'Info',
            parent=self.styles['Normal'],
            fontSize=13,
            fontName=self.chinese_font,
            alignment=TA_CENTER,
            leading=12
        )
        
        # 座位号样式（竖排）
        self.seat_style = ParagraphStyle(
            'Seat',
            parent=self.styles['Normal'],
            fontSize=8,
            fontName=self.chinese_font,
            alignment=TA_CENTER,
            leading=11,
            spaceBefore=0,
            spaceAfter=0
        )
        
        # 姓名样式（竖排）
        self.name_style = ParagraphStyle(
            'Name',
            parent=self.styles['Normal'],
            fontSize=7,
            fontName=self.chinese_font,
            alignment=TA_CENTER,
            leading=9,
            spaceBefore=0,
            spaceAfter=0
        )
        
        # 身份证号样式
        self.id_style = ParagraphStyle(
            'ID',
            parent=self.styles['Normal'],
            fontSize=5,
            fontName=self.chinese_font,
            alignment=TA_LEFT,
            leading=8
        )
        
        # 签名样式
        self.sign_style = ParagraphStyle(
            'Sign',
            parent=self.styles['Normal'],
            fontSize=6,
            fontName=self.chinese_font,
            alignment=TA_LEFT,
            leading=8
        )
        
        # 注意事项样式
        self.note_style = ParagraphStyle(
            'Note',
            parent=self.styles['Normal'],
            fontSize=7,
            fontName=self.chinese_font,
            alignment=TA_LEFT,
            leading=10,
            spaceBefore=5
        )
    
    def create_photo_image(
        self, 
        photo_path: Optional[str], 
        width: float = None, 
        height: float = None
    ) -> object:
        """
        创建照片图像
        
        Args:
            photo_path: 照片路径
            width: 宽度
            height: 高度
            
        Returns:
            图像对象或占位符表格
        """
        if width is None:
            width = self.photo_width
        if height is None:
            height = self.photo_height
        
        if photo_path and os.path.exists(photo_path):
            try:
                img = Image(photo_path, width=width, height=height)
                return img
            except Exception as e:
                logger.error(f"照片加载失败 {photo_path}: {e}")
        
        # 创建占位符
        data = [['照片']]
        table = Table(data, colWidths=[width], rowHeights=[height])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), self.chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
        ]))
        return table
    
    def create_student_cell(self, student_row: pd.Series) -> object:
        """
        创建单个考生信息单元格
        
        Args:
            student_row: 考生数据行
            
        Returns:
            包含考生信息的表格对象
        """
        # 获取照片
        id_card = str(student_row['身份证号'])
        photo_path = self.photo_map.get(id_card, None)
        photo_obj = self.create_photo_image(photo_path, self.photo_width, self.photo_height)
        
        # 座位号（竖排，顶部与照片平齐）
        seat_num = int(student_row['座位号'])
        seat_para = Paragraph(f"{seat_num:02d}", self.seat_style)
        
        # 姓名（竖排，独立区域，与座位号不重叠）
        name = str(student_row['姓名'])
        name_text = create_vertical_text(name)
        name_para = Paragraph(name_text, self.name_style)
        
        # 身份证号
        id_masked = str(id_card)
        
        # 右侧照片和下方信息
        right_data = [
            [photo_obj],
            [Paragraph(f"{id_masked}", self.id_style)],
            [Paragraph(f"我已知晓考试要求", self.sign_style)],
            [Paragraph(f"", self.sign_style)],
            [Paragraph(f"签名____________", self.sign_style)]
        ]
        right_table = Table(
            right_data,
            colWidths=[self.photo_width],
            rowHeights=[self.photo_height, 0.3*cm, 0.25*cm, 0.3*cm, 0.35*cm]
        )
        right_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))
        
        # 左侧竖排座位号和姓名（两个独立行，顶部与照片平齐）
        side_data = [
            [seat_para],  # 座位号（第一行，与照片顶部平齐）
            [name_para],  # 姓名（第二行，独立区域不重叠）
            [Spacer(1, self.photo_height - 0.8*cm)],  # 空白填充，保持与照片同高
        ]
        side_table = Table(
            side_data,
            colWidths=[self.side_width],
            rowHeights=[0.4*cm, 0.4*cm, None]  # 座位号和姓名各占固定高度
        )
        side_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 顶部对齐，与照片平齐
            ('LEFTPADDING', (0, 0), (-1, -1), 1),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        
        # 合并左右两侧
        combined_data = [[side_table, right_table]]
        combined_table = Table(
            combined_data,
            colWidths=[self.side_width, self.photo_width],
            rowHeights=[self.photo_height + 1.2*cm]
        )
        combined_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        
        return combined_table
    
    def create_header(
        self, 
        exam_info: Dict[str, str], 
        page_num: int, 
        total_pages: int
    ) -> List[List[object]]:
        """
        创建页眉
        
        Args:
            exam_info: 考试信息字典
            page_num: 当前页码
            total_pages: 总页数
            
        Returns:
            页眉表格数据
        """
        title = Paragraph(
            "<b>XX学校2026年XX考试图像信息核对单</b>", 
            self.title_style
        )
        
        info_text = (
            f"考点:{exam_info.get('exam_site', '')}&nbsp;&nbsp;&nbsp;&nbsp;"
            f"考试时间:{exam_info.get('date', '')}&nbsp;&nbsp;&nbsp;&nbsp;"
            f"考场:{exam_info.get('exam_room', '')}"
        )
        info = Paragraph(info_text, self.info_style)
        
        return [[title], [info]]
    
    def create_footer(self, exam_info: Dict[str, str]) -> List[object]:
        """
        创建页脚
        
        Args:
            exam_info: 考试信息字典
            
        Returns:
            页脚元素列表
        """
        notes = [
            Paragraph("1.考前20分钟，监考员持《考场图像单》核对考生身份，并指导参考考生在签名处签名。", self.note_style),
            Paragraph("2.考后30分钟，监考员用黑色签字笔将缺考考生姓名下方黑框涂黑。", self.note_style),
            Paragraph("3.课程代码前面的▲代表该课程可以使用无存储功能的计算器。", self.note_style),
            Paragraph("4.课程(14169)设计基础可携带铅笔、勾线笔、水粉、水彩、彩铅等绘图工具。；(04851)产品设计程序与方法可携带马克笔、彩铅、色粉等绘图工具。", self.note_style),
            Paragraph("5.诚信誓词抄写要求见《考试通知单》-《考场规则》第6条。", self.note_style),
            Paragraph("监考员签名：1._____________  2._____________", self.info_style),
        ]
        return notes
    
    def generate_room_pdf(
        self, 
        room_num: int, 
        exam_info: Dict[str, str]
    ) -> str:
        """
        生成单个考场的图像单（一页纸打印）
        
        Args:
            room_num: 考场号
            exam_info: 考试信息
            
        Returns:
            生成的PDF文件路径
        """
        room_data = self.exam_rooms[room_num]
        students = room_data.to_dict('records')
        
        # 输出文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"考场图像单_{room_num:02d}考场_{timestamp}.pdf"
        output_path = self.output_dir / output_filename
        
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            leftMargin=1.0*cm,
            rightMargin=1.0*cm,
            topMargin=1.0*cm,
            bottomMargin=0.8*cm
        )
        
        story = []
        
        # 页眉
        header_data = self.create_header(exam_info, 1, 1)
        header_table = Table(header_data, colWidths=[A4[0] - 2.4*cm])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 0.3*cm))
        
        # 考生表格（5 行×6 列=30 人，一页打印）
        student_table_data = []
        
        # 计算每列宽度（6 列）
        cell_total_width = (A4[0] - 2.4*cm - 0.5*cm * (self.cols - 1)) / self.cols
        photo_width = cell_total_width * 0.85  # 照片占列宽的 85%
        
        for row_idx in range(self.rows):
            row_data = []
            for col_idx in range(self.cols):
                idx = row_idx * self.cols + col_idx
                if idx < len(students):
                    student_cell = self.create_student_cell(students[idx])
                    row_data.append(student_cell)
                else:
                    row_data.append(Spacer(1, 0.5*cm))
            student_table_data.append(row_data)
        
        student_table = Table(
            student_table_data,
            colWidths=[cell_total_width] * self.cols,
            rowHeights=[self.photo_height + 1.2*cm] * self.rows
        )
        student_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))
        
        story.append(student_table)
        story.append(Spacer(1, 0.5*cm))
        
        # 页脚（注意事项）
        footer_notes = self.create_footer(exam_info)
        for note in footer_notes:
            story.append(note)
        
        story.append(Spacer(1, 0.5*cm))
        
        # 构建 PDF
        doc.build(story)
        
        logger.info(f"✅ 考场{room_num:02d}图像单生成完成: {output_path}")
        return str(output_path)
    
    def generate_all_rooms(
        self, 
        exam_date: str = '2026-04-1上午', 
        exam_site: str = '2幢教学楼'
    ) -> List[str]:
        """
        生成所有考场的图像单
        
        Args:
            exam_date: 考试日期
            exam_site: 考点
            
        Returns:
            生成的文件路径列表
        """
        generated_files = []
        
        exam_info = {
            'date': exam_date,
            'exam_site': exam_site
        }
        
        for room_num in sorted(self.exam_rooms.keys()):
            exam_info['exam_room'] = str(room_num)
            output_path = self.generate_room_pdf(room_num, exam_info)
            generated_files.append(output_path)
        
        logger.info(f"✅ 全部完成！共生成 {len(generated_files)} 个考场图像单")
        logger.info(f"📁 文件位置：{self.output_dir}")
        
        return generated_files

def main():
    # 数据路径 - 可通过命令行参数或配置文件指定
    data_dir = input("请输入数据目录路径 (默认: ./data): ").strip() or "./data"
    photo_dir = input("请输入照片目录路径 (默认: ./photos): ").strip() or "./photos"
    output_dir = input("请输入输出目录路径 (默认: ./output): ").strip() or "./output"
    
    try:
        # 创建生成器
        generator = ExamImageGenerator(data_dir, photo_dir, output_dir)
        
        print(f"\n开始生成考场图像单...")
        print(f"输出目录：{generator.output_dir}\n")
        
        # 生成所有考场
        generated_files = generator.generate_all_rooms(
            exam_date='2026-04-01上午09:00-11:30',
            exam_site='2幢教学楼'
        )
        
        print(f"\n{'='*60}")
        print(f"✅ 全部完成！共生成 {len(generated_files)} 个考场图像单")
        print(f"📁 文件位置：{generator.output_dir}")
        
    except FileNotFoundError as e:
        logger.error(f"❌ 文件或目录不存在: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ 生成过程中发生错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
