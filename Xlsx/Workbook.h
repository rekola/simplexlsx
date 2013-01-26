#ifndef __WORKBOOK_H__
#define __WORKBOOK_H__

#include "Worksheet.h"
#include "Chartsheet.h"
#include "SimpleXlsxDef.h"

namespace xmlw {
	class XmlStream;
}

namespace SimpleXlsx {
bool MakeDirectory(const _tstring& dirName);

/// @brief  This structure represents a handle to manage newly adding styles to avoid dublicating
struct StyleList {
private:
	std::vector<Border> borders;	///< borders set of values represent styled borders
	std::vector<Font> fonts;	///< fonts set of fonts to be declared
	std::vector<Fill> fills;	///< fills set of fills to be declared
	std::vector< std::pair<std::pair<int32_t, int32_t>, int32_t> > styleIndexes;///< styleIndexes vector of a number triplet contains links to style parts:
																				///         first - border id in borders
																				///         second - font id in fonts
																				///         third - fill id in fills
	std::vector< std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> > stylePos;///< stylePos vector of a number triplet contains style`s alignments and wrap sign:
																				///         first - EAlignHoriz value
																				///         second - EAlignVert value
																				///         third - wrap boolean value
public:
	StyleList() {
		borders.clear();
		fonts.clear();
		fills.clear();
		styleIndexes.clear();
		stylePos.clear();
	}

	/// @brief	For internal use (at the book saving)
	std::vector<Border> GetBorders()const { return borders; }
	/// @brief	For internal use (at the book saving)
	std::vector<Font> GetFonts()	const { return fonts; }
	/// @brief	For internal use (at the book saving)
	std::vector<Fill> GetFills()	const { return fills; }
	/// @brief	For internal use (at the book saving)
	std::vector< std::pair<std::pair<int32_t, int32_t>, int32_t> > GetIndexes() const { return styleIndexes; }
	/// @brief	For internal use (at the book saving)
	std::vector< std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> > GetPositions() const { return stylePos; }

	/// @brief  Adds a new style into collection if it is not exists yet
	/// @param  style Reference to the Style structure object
	/// @return Style index that should be used at data appending to a data sheet
	/// @note   If returned value is 0 - this is a default normal style and it is optional
	///         whether is will be added into column description or not
	///         (but better not to add to reduce size and resource consumption)
	int32_t Add(const Style& style) {
		std::pair<std::pair<int32_t, int32_t>, int32_t> styleLinks;

		bool addItem = true;
		for (int32_t i = 0; i < (int32_t)borders.size(); i++) {
			if (borders[i] == style.border) {
				addItem = false;
				styleLinks.first.first = i;
				break;
			}
		}

		if (addItem) {
			borders.push_back(style.border);
			styleLinks.first.first = borders.size() - 1;
		}

		addItem = true;
		for (int32_t i = 0; i < (int32_t)fonts.size(); i++) {
			if (fonts[i] == style.font) {
				addItem = false;
				styleLinks.first.second = i;
				break;
			}
		}

		if (addItem) {
			fonts.push_back(style.font);
			styleLinks.first.second = fonts.size() - 1;
		}

		addItem = true;
		for (int32_t i = 0; i < (int32_t)fills.size(); i++) {
			if (fills[i] == style.fill) {
				addItem = false;
				styleLinks.second = i;
				break;
			}
		}

		if (addItem) {
			fills.push_back(style.fill);
			styleLinks.second = fills.size() - 1;
		}

		for (int32_t i = 0; i < (int32_t)styleIndexes.size(); i++) {
			if (styleIndexes[i] == styleLinks &&
				stylePos[i].first.first == style.horizAlign &&
				stylePos[i].first.second == style.vertAlign &&
				stylePos[i].second == style.wrapText)
				return i;
		}

		std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> pos;
		pos.first.first = style.horizAlign;
		pos.first.second = style.vertAlign;
		pos.second = style.wrapText;
		stylePos.push_back(pos);

		styleIndexes.push_back(styleLinks);
		return styleIndexes.size() - 1;
	}
};

// ****************************************************************************
/// @brief	The class CWorkbook is used to manage creation, population and saving .xlsx files
// ****************************************************************************
class CWorkbook {
	_tstring                    m_temp_path;		///< path to the temporary directory (unique for a book)
	std::vector<_tstring>    	m_contentFiles;		///< a series of relative file pathes to be saved inside xlsx archive
	std::vector<CWorksheet*>    m_worksheets;		///< a series of data sheets
	std::vector<CChartsheet*>   m_charts;			///< a series of chart sheets
	std::map<_tstring, uint64_t>m_sharedStrings;	///<
	std::vector<Comment>		m_comments;			///<

public:
	// @section    Constructors / destructor
	CWorkbook();
	virtual ~CWorkbook();

	// @section    User interface
	StyleList m_styleList;

	CWorksheet& AddSheet(const _tstring& title);
	CWorksheet& AddSheet(const _tstring& title, std::vector<ColumnWidth>& colWidths);
	CWorksheet& AddSheet(const _tstring& title, uint32_t frozenWidth, uint32_t frozenHeight);
	CWorksheet& AddSheet(const _tstring& title, uint32_t frozenWidth, uint32_t frozenHeight, std::vector<ColumnWidth>& colWidths);

	CChartsheet& AddChart(const _tstring& title);
	CChartsheet& AddChart(const _tstring& title, EChartTypes type);

	bool Save(const _tstring& name);

private:
	void Init();

	bool SaveCore() const;
	bool SaveContentType();
	bool SaveApp() const;
	bool SaveTheme() const;
	bool SaveStyles() const;
	bool SaveChain();
	bool SaveComments();
	bool SaveSharedStrings();
	bool SaveWorkbook() const;

	bool SaveCommentList(std::vector<Comment*> &comments);
	void AddComment(xmlw::XmlStream &xml_stream, const Comment &comment) const;
	void AddCommentDrawing(xmlw::XmlStream &xml_stream, const Comment &comment) const;
	void AddFonts(xmlw::XmlStream& stream) const;
	void AddFills(xmlw::XmlStream& stream) const;
	void AddBorders(xmlw::XmlStream& stream) const;
	void AddBorder(xmlw::XmlStream& stream, const TCHAR *borderName, Border::BorderItem border) const;

	void ClearTemp();
};

}	// namespace SimpleXlsx

#endif	// __WORKBOOK_H__
