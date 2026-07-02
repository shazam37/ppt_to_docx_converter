from .core_ditto import PPTToDocConverterDitto, ConversionError
from .core_visual import PPTToDocConverterVisual
from .core_clean import PPTToDocConverterClean
from .core_enhanced import PPTToDocConverterEnhanced

# Use ditto converter as default (maximum visual fidelity - exact fonts, colors, alignment)
PPTToDocConverter = PPTToDocConverterDitto

__all__ = ['PPTToDocConverter', 'PPTToDocConverterDitto', 'PPTToDocConverterVisual', 'PPTToDocConverterClean', 'PPTToDocConverterEnhanced', 'ConversionError']
