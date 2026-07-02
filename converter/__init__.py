from .core_visual import PPTToDocConverterVisual, ConversionError
from .core_clean import PPTToDocConverterClean
from .core_enhanced import PPTToDocConverterEnhanced

# Use visual converter as default (charts as images, not tables!)
PPTToDocConverter = PPTToDocConverterVisual

__all__ = ['PPTToDocConverter', 'PPTToDocConverterVisual', 'PPTToDocConverterClean', 'PPTToDocConverterEnhanced', 'ConversionError']
