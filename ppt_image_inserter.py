#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 이미지 자동 삽입 스크립트
89개 업체의 이미지를 PPT에 자동으로 삽입합니다.
"""

import os
import glob
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import openpyxl

class PPTImageInserter:
    def __init__(self, template_ppt_path, base_image_dir, output_ppt_path, excel_path=None):
        """
        Args:
            template_ppt_path: 템플릿 PPT 파일 경로
            base_image_dir: 이미지가 있는 기본 디렉토리 (예: C:/Users/user/Downloads/서울/중구)
            output_ppt_path: 출력 PPT 파일 경로
            excel_path: 엑셀 파일 경로 (업체 순서 정보)
        """
        self.template_ppt_path = template_ppt_path
        self.base_image_dir = base_image_dir
        self.output_ppt_path = output_ppt_path
        self.excel_path = excel_path
        
    def load_shop_order_from_excel(self):
        """엑셀 파일에서 업체 순서를 읽어옵니다."""
        if not self.excel_path or not os.path.exists(self.excel_path):
            print("엑셀 파일을 찾을 수 없습니다. 기본 정렬을 사용합니다.")
            return None
        
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            
            shop_order = []
            for row in range(2, ws.max_row + 1):
                shop_name = ws.cell(row, 3).value  # 매장명은 3번째 컬럼
                if shop_name:
                    shop_order.append(shop_name.strip())
            
            print(f"엑셀에서 {len(shop_order)}개 업체 순서 로드 완료")
            return shop_order
            
        except Exception as e:
            print(f"엑셀 파일 읽기 오류: {e}")
            return None
    
    def find_shop_directories(self):
        """업체 디렉토리 목록을 찾고 엑셀 순서대로 정렬합니다."""
        shop_dirs_dict = {}
        
        # 기본 경로에서 하위 디렉토리 탐색
        if os.path.exists(self.base_image_dir):
            for item in os.listdir(self.base_image_dir):
                item_path = os.path.join(self.base_image_dir, item)
                if os.path.isdir(item_path):
                    # 업체 폴더 내에 업체_xxx.jpg 파일이 있는지 확인
                    image_files = glob.glob(os.path.join(item_path, "업체", "업체_*.jpg"))
                    if image_files:
                        shop_dirs_dict[item] = {
                            'name': item,
                            'path': item_path,
                            'images': sorted(image_files)
                        }
        
        # 엑셀 파일에서 순서 로드
        shop_order = self.load_shop_order_from_excel()
        
        if shop_order:
            # 엑셀 순서대로 정렬
            ordered_shops = []
            for shop_name in shop_order:
                if shop_name in shop_dirs_dict:
                    ordered_shops.append(shop_dirs_dict[shop_name])
                else:
                    print(f"  경고: '{shop_name}' 폴더를 찾을 수 없습니다.")
            
            # 엑셀에 없는 업체는 마지막에 추가
            for shop_name, shop_info in sorted(shop_dirs_dict.items()):
                if shop_name not in shop_order:
                    print(f"  추가: '{shop_name}' (엑셀에 없는 업체)")
                    ordered_shops.append(shop_info)
            
            return ordered_shops
        else:
            # 엑셀이 없으면 알파벳 순서로 정렬
            return [shop_dirs_dict[key] for key in sorted(shop_dirs_dict.keys())]
    
    def get_image_dimensions(self, image_path):
        """이미지의 크기를 확인합니다."""
        try:
            with Image.open(image_path) as img:
                return img.size  # (width, height)
        except Exception as e:
            print(f"이미지 읽기 오류 ({image_path}): {e}")
            return None
    
    def add_shop_to_ppt(self, prs, shop_info, shop_number):
        """업체 정보를 PPT에 추가합니다."""
        images = shop_info['images']
        shop_name = shop_info['name']
        
        # 표지 슬라이드 추가 (업체명)
        slide_layout = prs.slide_layouts[5]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # 제목 텍스트 박스 추가
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(12.33)
        height = Inches(1)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.text = f"{shop_name}"
        
        # 텍스트 스타일 설정
        for paragraph in text_frame.paragraphs:
            paragraph.font.size = Inches(0.5)
            paragraph.font.bold = True
        
        print(f"  - 표지 슬라이드 추가: {shop_name}")
        
        # 이미지 슬라이드 추가
        for idx, image_path in enumerate(images):
            # 슬라이드 종류 결정
            if "가격표" in os.path.basename(image_path) or idx == 0:
                slide_title = "가격표"
            else:
                slide_title = "인테리어"
            
            slide = prs.slides.add_slide(slide_layout)
            
            # 제목 추가
            textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(1))
            text_frame = textbox.text_frame
            text_frame.text = slide_title
            for paragraph in text_frame.paragraphs:
                paragraph.font.size = Inches(0.4)
                paragraph.font.bold = True
            
            # 이미지 추가
            try:
                # 이미지 크기 확인
                img_dims = self.get_image_dimensions(image_path)
                if img_dims:
                    img_width, img_height = img_dims
                    aspect_ratio = img_width / img_height
                    
                    # 슬라이드 중앙에 배치
                    max_width = Inches(11)
                    max_height = Inches(5.5)
                    
                    if aspect_ratio > max_width / max_height:
                        # 가로가 더 긴 경우
                        width = max_width
                        height = width / aspect_ratio
                    else:
                        # 세로가 더 긴 경우
                        height = max_height
                        width = height * aspect_ratio
                    
                    left = (prs.slide_width - width) / 2
                    top = Inches(1.5)
                    
                    slide.shapes.add_picture(image_path, left, top, width=width, height=height)
                    print(f"  - 이미지 추가: {os.path.basename(image_path)}")
                else:
                    print(f"  - 이미지 크기 확인 실패: {image_path}")
                    
            except Exception as e:
                print(f"  - 이미지 추가 실패 ({image_path}): {e}")
    
    def create_ppt(self):
        """전체 PPT를 생성합니다."""
        print("=" * 60)
        print("PPT 이미지 자동 삽입 시작")
        print("=" * 60)
        
        # 업체 디렉토리 찾기
        print(f"\n이미지 디렉토리 스캔 중: {self.base_image_dir}")
        shop_dirs = self.find_shop_directories()
        
        if not shop_dirs:
            print(f"오류: 업체 디렉토리를 찾을 수 없습니다.")
            print(f"경로를 확인하세요: {self.base_image_dir}")
            return False
        
        print(f"발견된 업체 수: {len(shop_dirs)}")
        
        # 템플릿 PPT 로드
        print(f"\n템플릿 PPT 로드 중: {self.template_ppt_path}")
        prs = Presentation(self.template_ppt_path)
        
        # 표지 슬라이드 추가
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        textbox = slide.shapes.add_textbox(Inches(4), Inches(3), Inches(5.33), Inches(1.5))
        text_frame = textbox.text_frame
        text_frame.text = "세신샵 업체 정보"
        for paragraph in text_frame.paragraphs:
            paragraph.font.size = Inches(0.6)
            paragraph.font.bold = True
        
        print("\n업체별 슬라이드 생성 중...")
        
        # 각 업체별로 슬라이드 추가
        for idx, shop_info in enumerate(shop_dirs, 1):
            print(f"\n[{idx}/{len(shop_dirs)}] {shop_info['name']} 처리 중...")
            print(f"  이미지 수: {len(shop_info['images'])}")
            self.add_shop_to_ppt(prs, shop_info, idx)
        
        # PPT 저장
        print(f"\nPPT 저장 중: {self.output_ppt_path}")
        prs.save(self.output_ppt_path)
        
        print("=" * 60)
        print(f"완료! 총 {len(shop_dirs)}개 업체의 슬라이드가 생성되었습니다.")
        print(f"출력 파일: {self.output_ppt_path}")
        print("=" * 60)
        
        return True


def main():
    """메인 실행 함수"""
    
    # ========================================
    # 설정 값 (필요에 따라 수정하세요)
    # ========================================
    
    # Windows 경로 예시 (실제 경로로 변경하세요)
    BASE_IMAGE_DIR = r"C:\Users\user\Downloads\서울\중구"
    
    # Linux/Mac 경로 예시
    # BASE_IMAGE_DIR = "/home/user/Downloads/서울/중구"
    
    # 템플릿 PPT 파일
    TEMPLATE_PPT = "/home/user/uploaded_files/세신샵.pptx"
    
    # 엑셀 파일 (업체 순서 정보)
    EXCEL_FILE = "/home/user/uploaded_files/리스트_네이버지도링크추가 - 복사본.xlsx"
    
    # 출력 PPT 파일
    OUTPUT_PPT = "/home/user/webapp/세신샵_완성본.pptx"
    
    # ========================================
    
    # 경로 확인
    if not os.path.exists(TEMPLATE_PPT):
        print(f"오류: 템플릿 PPT 파일을 찾을 수 없습니다: {TEMPLATE_PPT}")
        return
    
    # PPT 생성기 실행
    inserter = PPTImageInserter(
        template_ppt_path=TEMPLATE_PPT,
        base_image_dir=BASE_IMAGE_DIR,
        output_ppt_path=OUTPUT_PPT,
        excel_path=EXCEL_FILE
    )
    
    inserter.create_ppt()


if __name__ == "__main__":
    main()
