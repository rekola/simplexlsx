#ifndef __SIMPLE_XLSX_DEF_H__
#define __SIMPLE_XLSX_DEF_H__

#include <stdint.h>
#include <vector>
#include <map>
#include <string>
#include <utility>
#include <ctime>
#include <fstream>

#ifdef _WIN32
#include <tchar.h>
#else
#include "../tchar.h"
#endif  // _WIN32

#ifdef _UNICODE
	typedef std::wofstream 		_tofstream;
	typedef std::wstring		_tstring;
	typedef std::wostringstream	_tstringstream;
	typedef std::wostream		_tstream;
#else
	typedef std::ofstream		_tofstream;
	typedef std::string			_tstring;
	typedef std::ostringstream	_tstringstream;
	typedef std::ostream		_tstream;
#endif // _UNICODE

#define SIMPLE_XLSX_VERSION	_T("0.18")

namespace SimpleXlsx {

/// @brief	Possible chart types
enum EChartTypes {
	CHART_NONE = -1,
	CHART_LINEAR = 0,
	CHART_BAR,
	CHART_SCATTER,
};

/// @brief  Possible border attributes
enum EBorderStyle {
	BORDER_NONE = 0,
	BORDER_THIN,
	BORDER_MEDIUM,
	BORDER_DASHED,
	BORDER_DOTTED,
	BORDER_THICK,
	BORDER_DOUBLE,
	BORDER_HAIR,
	BORDER_MEDIUM_DASHED,
	BORDER_DASH_DOT,
	BORDER_MEDIUM_DASH_DOT,
	BORDER_DASH_DOT_DOT,
	BORDER_MEDIUM_DASH_DOT_DOT,
	BORDER_SLANT_DASH_DOT
};

/// @brief  Additional font attributes
enum EFontAttributes {
	FONT_NORMAL     = 0,
	FONT_BOLD       = 1,
	FONT_ITALIC     = 2,
	FONT_UNDERLINED = 4,
	FONT_STRIKE     = 8,
	FONT_OUTLINE    = 16,
	FONT_SHADOW     = 32,
	FONT_CONDENSE   = 64,
	FONT_EXTEND     = 128,
};

/// @brief  Fill`s pattern type enumeration
enum EPatternType {
	PATTERN_NONE = 0,
	PATTERN_SOLID,
	PATTERN_MEDIUM_GRAY,
	PATTERN_DARK_GRAY,
	PATTERN_LIGHT_GRAY,
	PATTERN_DARK_HORIZ,
	PATTERN_DARK_VERT,
	PATTERN_DARK_DOWN,
	PATTERN_DARK_UP,
	PATTERN_DARK_GRID,
	PATTERN_DARK_TRELLIS,
	PATTERN_LIGHT_HORIZ,
	PATTERN_LIGHT_VERT,
	PATTERN_LIGHT_DOWN,
	PATTERN_LIGHT_UP,
	PATTERN_LIGHT_GRID,
	PATTERN_LIGHT_TRELLIS,
	PATTERN_GRAY_125,
	PATTERN_GRAY_0625
};

/// @brief	Text horizontal alignment enumerations
enum EAlignHoriz {
	ALIGN_H_NONE = 0,
	ALIGN_H_LEFT,
	ALIGN_H_CENTER,
	ALIGN_H_RIGHT
};

/// @brief	Text vertical alignment enumerations
enum EAlignVert {
	ALIGN_V_NONE = 0,
	ALIGN_V_TOP,
	ALIGN_V_CENTER,
	ALIGN_V_BOTTOM
};

/// @brief  Font describes a font that can be added into final document stylesheet
/// @see    EFontAttributes
struct Font {
	int32_t size;		///< size font size
	_tstring name;		///< name font name (there is no enumeration or preset values, it should be used carefully)
	bool theme;			///< theme if true then color is not taken into account
	_tstring color;		///< color color format: AARRGGBB - (alpha, red, green, blue). If empty default theme is used
	int32_t attributes;	///< combination of additinal font flags (EFontAttributes)

	Font() {
		Clear();
	}

	void Clear() {
		size = 11;
		name = _T("Calibri");
		theme = true;
		color = _T("");
		attributes = FONT_NORMAL;
	}

	bool operator==(const Font& _font) const {
		return (size == _font.size && name == _font.name && theme == _font.theme &&
				color == _font.color && attributes == _font.attributes);
	}
};

/// @brief  Fill describes a fill that can be added into final document stylesheet
/// @note   Current version describes the pattern fill only
/// @see    EPatternType
struct Fill {
	EPatternType patternType;	///< patternType
	_tstring fgColor;			///< fgColor foreground color format: AARRGGBB - (alpha, red, green, blue). Can be left unset
	_tstring bgColor;			///< bgColor background color format: AARRGGBB - (alpha, red, green, blue). Can be left unset

	Fill() {
		Clear();
	}

	void Clear() {
		patternType = PATTERN_NONE;
		fgColor = _T("");
		bgColor = _T("");
	}

	bool operator==(const Fill& _fill) const {
		return (patternType == _fill.patternType && fgColor == _fill.fgColor && bgColor == _fill.bgColor);
	}
};

/// @brief  Border describes a border style that can be added into final document stylesheet
/// @see    EBorderStyle
struct Border {
	/// @brief	BorderItem describes border items (e.g. left, right, bottom, top, diagonal sides)
	struct BorderItem {
		EBorderStyle style;		///< style border style
		_tstring color;			///< colour border colour format: AARRGGBB - (alpha, red, green, blue). Can be left unset

		BorderItem() {
			Clear();
		}

		void Clear() {
			style = BORDER_NONE;
			color = _T("");
		}

		bool operator==(const BorderItem& _borderItem) const {
			return (color == _borderItem.color && style == _borderItem.style);
		}
	};

	BorderItem left;			///< left left side style
	BorderItem right;			///< right right side style
	BorderItem bottom;			///< bottom bottom side style
	BorderItem top;				///< top top side style
	BorderItem diagonal;		///< diagonal diagonal side style
	bool isDiagonalUp;			///< isDiagonalUp indicates whether this diagonal border should be used
	bool isDiagonalDown;		///< isDiagonalDown indicates whether this diagonal border should be used

	Border() {
		Clear();
	}

	void Clear() {
		isDiagonalUp = false;
		isDiagonalDown = false;

		left.Clear();
		right.Clear();
		bottom.Clear();
		top.Clear();
	}

	bool operator==(const Border& _border) const {
		return (isDiagonalUp == _border.isDiagonalUp && isDiagonalDown == _border.isDiagonalDown &&
				left == _border.left && right == _border.right &&
				bottom == _border.bottom && top == _border.top);
	}
};

/// @brief  Style describes a set of styling parameter that can be used into final document
/// @see    EBorder
/// @see    EAlignHoriz
/// @see    EAlignVert
struct Style {
	Font font;				///< font structure object describes font
	Fill fill;				///< fill structure object describes fill
	Border border;			///< border combination of border attributes
	EAlignHoriz horizAlign;	///< horizAlign cell content horizontal alignment value
	EAlignVert vertAlign;	///< vertAlign cell content vertical alignment value
	bool wrapText;			///< wrapText text wrapping property

	Style() {
		Clear();
	}

	void Clear() {
		font.Clear();
		fill.Clear();
		border.Clear();
		horizAlign = ALIGN_H_NONE;
		vertAlign = ALIGN_V_NONE;
		wrapText = false;
	}
};

/// @brief	Cell coordinate structure
typedef struct _CellCoord {
	uint32_t row;	///< row (starts from 1)
	uint32_t col;	///< col (starts from 0)

	_CellCoord() { Clear(); }
	_CellCoord(uint32_t _r, uint32_t _c) : row(_r), col(_c) {}

	void Clear() {
		row = 1;
		col = 0;
	}
} CellCoord;

/// @brief	Column width structure
typedef struct _ColumnWidth {
	uint32_t colFrom;	///< column range from (starts from 0)
	uint32_t colTo;		///< column range to (starts from 0)
	float width;		///< specified width

	_ColumnWidth() {colFrom = colTo = 0; width = 15;}
	_ColumnWidth(uint32_t min, uint32_t max, float w) : colFrom(min), colTo(max), width(w) {}
} ColumnWidth;

typedef struct _CellDataStr {
    _tstring value;
    int32_t style_id;

    _CellDataStr() : value(_T("")), style_id(0) {}
} CellDataStr;	///< cell data:style pair
typedef struct _CellDataTime {
    time_t value;
    int32_t style_id;

    _CellDataTime() : value(0), style_id(0) {}
} CellDataTime;	///< cell data:style pair
typedef struct _CellDataInt {
    int32_t value;
    int32_t style_id;

    _CellDataInt() : value(0), style_id(0) {}
} CellDataInt;	///< cell data:style pair
typedef struct _CellDataUInt {
    uint32_t value;
    int32_t style_id;

    _CellDataUInt() : value(0), style_id(0) {}
} CellDataUInt;	///< cell data:style pair
typedef struct _CellDataDbl {
    double value;
    int32_t style_id;

    _CellDataDbl() : value(0.0), style_id(0) {}
} CellDataDbl;	///< cell data:style pair
typedef struct _CellDataFlt {
    float value;
    int32_t style_id;

    _CellDataFlt() : value(0.0), style_id(0) {}
} CellDataFlt;	///< cell data:style pair

/// @brief	This structure describes comment item that can added to a cell
struct Comment {
	int sheetIndex;										///< sheetIndex internal page index (must not be changed manually)
	std::vector<std::pair<Font, _tstring> > contents;	///< contents set of contents with specified fonts
	CellCoord cellRef;									///< cellRef reference to the cell
	_tstring fillColor;									///< fillColor comment box background colour (format: #RRGGBB)
	bool isHidden;										///< isHidden determines if comments box is hidden
	int x;												///< x absolute position in pt (can be left unset)
	int y;												///< y absolute position in pt (can be left unset)
	int width;											///< width width in pt (can be left unset)
	int height;											///< height height in pt (can be left unset)

	Comment() {
		Clear();
	}

	void Clear() {
		contents.clear();
		cellRef.Clear();
		fillColor = _T("#FFEFD5");	// papaya whip
		isHidden = true;
		x = y = 50;
		width = height = 100;
	}

	bool operator < (const Comment& _comm) const {
        return (sheetIndex < _comm.sheetIndex);
    }
};

}	// namespace SimpleXlsx

#endif // __SIMPLE_XLSX_DEF_H__