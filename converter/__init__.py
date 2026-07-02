from .core_enhanced import PPTToDocConverterEnhanced, ConversionError

# Use enhanced converter as default
PPTToDocConverter = PPTToDocConverterEnhanced

__all__ = ['PPTToDocConverter', 'PPTToDocConverterEnhanced', 'ConversionError']
