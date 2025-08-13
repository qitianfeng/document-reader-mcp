#!/usr/bin/env python3
"""
简单图表阅读器 - 直接读取图片文件并分析内容
无需复杂的OCR配置，基于图像结构分析
"""

import cv2
import numpy as np
from pathlib import Path
from typing import Dict, List, Any, Tuple
import json

class SimpleDiagramReader:
    """简单图表阅读器"""
    
    def __init__(self):
        self.diagram_patterns = {
            "flowchart": {
                "keywords": ["流程", "开始", "结束", "判断", "处理"],
                "shapes": ["rectangles", "diamonds", "ovals"],
                "connections": "arrows"
            },
            "sequence": {
                "keywords": ["时序", "调用", "返回", "请求", "响应"],
                "shapes": ["vertical_lines", "rectangles"],
                "connections": "horizontal_arrows"
            },
            "architecture": {
                "keywords": ["架构", "系统", "服务", "数据库", "接口"],
                "shapes": ["rectangles", "circles", "cylinders"],
                "connections": "bidirectional"
            }
        }
    
    def read_diagram_from_file(self, image_path: str) -> Dict[str, Any]:
        """直接从图片文件读取并分析图表内容"""
        try:
            # 读取图片
            img = cv2.imread(image_path)
            if img is None:
                return {"error": f"无法读取图片: {image_path}"}
            
            # 基本信息
            h, w, c = img.shape
            file_size = Path(image_path).stat().st_size
            
            result = {
                "file_info": {
                    "path": image_path,
                    "filename": Path(image_path).name,
                    "size": f"{file_size / 1024:.1f} KB",
                    "dimensions": f"{w}×{h}",
                    "channels": c
                },
                "analysis": self._analyze_diagram_structure(img),
                "interpretation": self._interpret_diagram_content(img, image_path)
            }
            
            return result
            
        except Exception as e:
            return {"error": f"分析图片时出错: {str(e)}"}
    
    def _analyze_diagram_structure(self, img: np.ndarray) -> Dict[str, Any]:
        """分析图表结构"""
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 边缘检测
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        
        # 检测基本形状
        rectangles = self._detect_rectangles(gray)
        circles = self._detect_circles(gray)
        lines = self._detect_lines(edges)
        
        # 分析布局
        layout = self._analyze_layout(gray)
        
        return {
            "shapes": {
                "rectangles": len(rectangles),
                "circles": len(circles),
                "lines": len(lines)
            },
            "complexity": self._calculate_complexity(rectangles, circles, lines),
            "layout": layout,
            "dominant_direction": self._get_dominant_direction(lines)
        }
    
    def _interpret_diagram_content(self, img: np.ndarray, image_path: str) -> Dict[str, Any]:
        """解释图表内容（基于结构特征推断）"""
        analysis = self._analyze_diagram_structure(img)
        
        # 根据文件名推断类型
        filename = Path(image_path).name.lower()
        type_hints = []
        
        if any(word in filename for word in ["flow", "流程", "process"]):
            type_hints.append("flowchart")
        elif any(word in filename for word in ["sequence", "时序", "seq"]):
            type_hints.append("sequence")
        elif any(word in filename for word in ["arch", "架构", "system"]):
            type_hints.append("architecture")
        
        # 根据结构特征推断
        shapes = analysis["shapes"]
        predicted_type = self._predict_diagram_type(shapes, analysis.get("layout", {}))
        
        # 生成内容描述
        content_description = self._generate_content_description(analysis, predicted_type)
        
        return {
            "predicted_type": predicted_type,
            "confidence": self._calculate_confidence(shapes, predicted_type),
            "type_hints_from_filename": type_hints,
            "content_description": content_description,
            "technical_elements": self._identify_technical_elements(analysis)
        }
    
    def _detect_rectangles(self, gray: np.ndarray) -> List[Dict]:
        """检测矩形"""
        # 边缘检测
        edges = cv2.Canny(gray, 50, 150)
        
        # 查找轮廓
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        rectangles = []
        for contour in contours:
            # 近似轮廓
            epsilon = 0.02 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)
            
            # 检查是否为矩形（4个顶点）
            if len(approx) == 4:
                area = cv2.contourArea(contour)
                if area > 500:  # 过滤小的噪声
                    x, y, w, h = cv2.boundingRect(contour)
                    rectangles.append({
                        "bbox": (int(x), int(y), int(w), int(h)),
                        "area": float(area),
                        "aspect_ratio": float(w / h) if h > 0 else 0.0
                    })
        
        return rectangles
    
    def _detect_circles(self, gray: np.ndarray) -> List[Dict]:
        """检测圆形"""
        circles = cv2.HoughCircles(
            gray, cv2.HOUGH_GRADIENT, dp=1, minDist=50,
            param1=50, param2=30, minRadius=15, maxRadius=150
        )
        
        circle_list = []
        if circles is not None:
            circles = np.round(circles[0, :]).astype("int")
            for (x, y, r) in circles:
                circle_list.append({
                    "center": (int(x), int(y)),
                    "radius": int(r),
                    "area": float(np.pi * r * r)
                })
        
        return circle_list
    
    def _detect_lines(self, edges: np.ndarray) -> List[Dict]:
        """检测直线"""
        lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=80, minLineLength=50, maxLineGap=10)
        
        line_list = []
        if lines is not None:
            for line in lines:
                x1, y1, x2, y2 = line[0]
                length = np.sqrt((x2-x1)**2 + (y2-y1)**2)
                angle = np.arctan2(y2-y1, x2-x1) * 180 / np.pi
                
                line_list.append({
                    "start": (int(x1), int(y1)),
                    "end": (int(x2), int(y2)),
                    "length": float(length),
                    "angle": float(angle),
                    "direction": self._classify_line_direction(angle)
                })
        
        return line_list
    
    def _classify_line_direction(self, angle: float) -> str:
        """分类线条方向"""
        angle = abs(angle)
        if angle < 30 or angle > 150:
            return "horizontal"
        elif 60 < angle < 120:
            return "vertical"
        else:
            return "diagonal"
    
    def _analyze_layout(self, gray: np.ndarray) -> Dict[str, Any]:
        """分析布局"""
        h, w = gray.shape
        
        # 分析密度分布
        top_half = gray[:h//2, :]
        bottom_half = gray[h//2:, :]
        left_half = gray[:, :w//2]
        right_half = gray[:, w//2:]
        
        # 计算各区域的边缘密度
        top_edges = cv2.Canny(top_half, 50, 150).sum()
        bottom_edges = cv2.Canny(bottom_half, 50, 150).sum()
        left_edges = cv2.Canny(left_half, 50, 150).sum()
        right_edges = cv2.Canny(right_half, 50, 150).sum()
        
        return {
            "aspect_ratio": float(w / h),
            "density_distribution": {
                "top": int(top_edges),
                "bottom": int(bottom_edges),
                "left": int(left_edges),
                "right": int(right_edges)
            },
            "primary_orientation": "landscape" if w > h else "portrait"
        }
    
    def _get_dominant_direction(self, lines: List[Dict]) -> str:
        """获取主要方向"""
        if not lines:
            return "unknown"
        
        horizontal = sum(1 for line in lines if line["direction"] == "horizontal")
        vertical = sum(1 for line in lines if line["direction"] == "vertical")
        diagonal = sum(1 for line in lines if line["direction"] == "diagonal")
        
        if horizontal > vertical and horizontal > diagonal:
            return "horizontal"
        elif vertical > horizontal and vertical > diagonal:
            return "vertical"
        else:
            return "mixed"
    
    def _calculate_complexity(self, rectangles: List, circles: List, lines: List) -> int:
        """计算复杂度"""
        return len(rectangles) * 2 + len(circles) * 2 + len(lines)
    
    def _predict_diagram_type(self, shapes: Dict, layout: Dict) -> str:
        """预测图表类型"""
        rect_count = shapes.get("rectangles", 0)
        circle_count = shapes.get("circles", 0)
        line_count = shapes.get("lines", 0)
        
        # 基于形状特征判断
        if rect_count > 5 and line_count > 10:
            if layout.get("primary_orientation") == "landscape":
                return "流程图 (Flowchart)"
            else:
                return "组织架构图 (Organizational Chart)"
        elif circle_count > 3 and line_count > 5:
            return "网络图 (Network Diagram)"
        elif rect_count > 2 and line_count > 8:
            return "系统架构图 (System Architecture)"
        elif line_count > 20:
            return "时序图 (Sequence Diagram)"
        else:
            return "技术图表 (Technical Diagram)"
    
    def _calculate_confidence(self, shapes: Dict, predicted_type: str) -> float:
        """计算置信度"""
        base_confidence = 0.6
        
        rect_count = shapes.get("rectangles", 0)
        circle_count = shapes.get("circles", 0)
        line_count = shapes.get("lines", 0)
        
        if "流程图" in predicted_type and rect_count > 3 and line_count > 5:
            return 0.85
        elif "架构图" in predicted_type and (rect_count > 2 or circle_count > 2):
            return 0.80
        elif "时序图" in predicted_type and line_count > 15:
            return 0.75
        
        return base_confidence
    
    def _generate_content_description(self, analysis: Dict, diagram_type: str) -> str:
        """生成内容描述"""
        shapes = analysis["shapes"]
        complexity = analysis["complexity"]
        
        description = f"这是一个{diagram_type}，"
        
        if complexity > 50:
            description += "结构较为复杂，"
        elif complexity > 20:
            description += "结构中等复杂，"
        else:
            description += "结构相对简单，"
        
        description += f"包含{shapes['rectangles']}个矩形框、{shapes['circles']}个圆形和{shapes['lines']}条连接线。"
        
        # 根据类型添加特定描述
        if "流程图" in diagram_type:
            description += "显示了业务流程的各个步骤和决策点。"
        elif "架构图" in diagram_type:
            description += "展示了系统组件之间的关系和交互。"
        elif "时序图" in diagram_type:
            description += "描述了不同对象间的时间序列交互。"
        
        return description
    
    def _identify_technical_elements(self, analysis: Dict) -> List[str]:
        """识别技术元素"""
        elements = []
        shapes = analysis["shapes"]
        
        if shapes["rectangles"] > 0:
            elements.append("处理节点")
        if shapes["circles"] > 0:
            elements.append("状态节点")
        if shapes["lines"] > 5:
            elements.append("流程连接")
        if analysis["complexity"] > 30:
            elements.append("复杂逻辑")
        
        return elements

def analyze_all_extracted_images() -> Dict[str, Any]:
    """分析所有提取的图片"""
    reader = SimpleDiagramReader()
    results = {}
    
    extracted_dir = Path("extracted_images")
    if not extracted_dir.exists():
        return {"error": "未找到 extracted_images 文件夹"}
    
    image_files = list(extracted_dir.glob("*.png")) + list(extracted_dir.glob("*.jpg"))
    
    for image_file in image_files:
        print(f"📊 分析图片: {image_file.name}")
        result = reader.read_diagram_from_file(str(image_file))
        results[image_file.name] = result
    
    return results

def analyze_single_image(image_path: str) -> Dict[str, Any]:
    """分析单个图片"""
    reader = SimpleDiagramReader()
    return reader.read_diagram_from_file(image_path)

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        # 分析指定图片
        result = analyze_single_image(sys.argv[1])
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        # 分析所有图片
        results = analyze_all_extracted_images()
        
        print("🎯 图表分析汇总:")
        for filename, result in results.items():
            if "error" not in result:
                interpretation = result.get("interpretation", {})
                print(f"\n📋 {filename}:")
                print(f"  类型: {interpretation.get('predicted_type', '未知')}")
                print(f"  置信度: {interpretation.get('confidence', 0):.1%}")
                print(f"  描述: {interpretation.get('content_description', '无描述')}")