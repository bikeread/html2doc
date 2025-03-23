"""
转换器工厂，用于创建HTML转DOCX转换器实例
采用工厂方法模式，便于未来扩展新的转换器类型
"""
from converters.base_converter import BaseConverter
from converters.docx_converter import DocxConverter


class ConverterFactory:
    """转换器工厂类"""
    
    @staticmethod
    def get_converter(converter_type: str) -> BaseConverter:
        """
        根据类型获取对应的转换器实例
        
        Args:
            converter_type: 转换器类型，目前只支持'docx'
            
        Returns:
            转换器实例
            
        Raises:
            ValueError: 当指定了不支持的转换器类型时
        """
        if converter_type.lower() == 'docx':
            return DocxConverter()
        else:
            raise ValueError(f"不支持的转换器类型: {converter_type}，目前仅支持'docx'") 