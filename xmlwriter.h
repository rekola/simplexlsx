#ifndef __XMLWRITER_H_2CC1D410_9DE8_4452_B8F0_F987EF152E06
#define __XMLWRITER_H_2CC1D410_9DE8_4452_B8F0_F987EF152E06

#include	<stack>
#include	<sstream>

#include	"Xlsx/SimpleXlsxDef.h"

namespace xmlw {

class XmlStream {
public:

	// XML version constants
	enum {versionMajor = 1, versionMinor = 0};

	// Internal helper class
	struct Controller {
		typedef	enum {whatProlog, whatTag, whatTagEnd, whatAttribute, whatCharData}	what_type;

		what_type	what;
		_tstring str;

		inline Controller(const Controller& c) : what(c.what), str(c.str) {}
		inline Controller(const what_type _what) : what(_what){}

		// use template constructor because string field <str> may be initialized
		// from different sources: TCHAR*, _tstring etc
		template<class t>
		inline Controller(const what_type _what, const t& _str) : what(_what), str(_str) {}
	};

	// XmlStream refers _tstream object to perform actual output operations
	inline XmlStream(_tstream&	_s) : state(stateNone), s(_s), prologWritten(false) {
//#ifndef _WIN32
        // in linux locale is not C default, so integer number grouping is on
        // back it into regular way by setting "C" locale
        s.imbue(std::locale("C"));
//#endif  // _WIN32
	}

	// Before destroying check whether all the open tags are closed
	~XmlStream() {
		if (stateTagName == state) {
			s << _T("/>");
			state = stateNone;
		}
		while (tags.size())
			endTag(tags.top());
	}

	// default behaviour - delegate object output to std::stream
	template<class t>
	XmlStream& operator<<(const t& value) {
		if (stateTagName == state)
			tagName << value;
		s << value;
		return *this;
	}

	// this is the main working horse
	// and it's long a little
	XmlStream& operator<<(const Controller& controller) {

		switch (controller.what) {
		case Controller::whatTag:
			closeTagStart();
			s << _T('<');
			if (controller.str.empty()) {
				clearTagName();
				state = stateTagName;
			}
			else {
				s << controller.str;
				tags.push(controller.str);
				state = stateTag;
			}
			break;	//	Controller::whatTag

		case Controller::whatTagEnd:
			endTag(controller.str);
			break;	//	Controller::whatTagEnd

		case Controller::whatAttribute:
			switch (state) {
			case stateTagName:
				tags.push(tagName.str());
				break;

			case stateAttribute:
				s << _T('\"');
            default:
                break;
			}

			if (stateNone != state) {
				s << _T(' ') << controller.str << _T("=\"");
				state = stateAttribute;
			}
			// else throw some error - unexpected attribute (out of any tag)

			break;	//	Controller::whatAttribute

		case Controller::whatCharData:
			closeTagStart();
			state = stateNone;
			break;	//	Controller::whatCharData

		case Controller::whatProlog:
			if (!prologWritten && stateNone == state) {
				s << _T("<?xml version=\"") << versionMajor << _T('.') << versionMinor << _T("\" encoding=\"UTF-8\" standalone=\"yes\" ?>\n");
				prologWritten = true;
			}
			break;	//	Controller::whatProlog
		}

		return	*this;
	}

private:
	// state of the stream
	typedef	enum {stateNone, stateTag, stateAttribute, stateTagName}	state_type;

	// tag name stack
	typedef std::stack<_tstring>	tag_stack_type;

	tag_stack_type	tags;
	state_type	state;
	_tstream&	s;
	bool	prologWritten;
	_tstringstream	tagName;

	// I don't know any way easier (legal) to clear std::stringstream...
	inline void clearTagName() {
		const _tstring	empty_str;
		tagName.rdbuf()->str(empty_str);
	}

	// Close current tag
	void closeTagStart(bool self_closed = false) {
		if (stateTagName == state)
			tags.push(tagName.str());

		// note: absence of 'break's is not an error
		switch (state) {
		case stateAttribute:
			s << _T('\"');

		case stateTagName:
		case stateTag:
			if (self_closed)
				s << _T('/');
			s << _T('>');
        default:
            break;
		}
	}

	// Close tag (may be with closing all of its children)
	void endTag(const _tstring& tag) {
		bool	brk = false;

		while (tags.size() > 0 && !brk) {
			if (stateNone == state) {
				s << _T("</") << tags.top() << _T('>');
			}
			else {
				closeTagStart(true);
				state = stateNone;
			}
			brk = tag.empty() || tag == tags.top();
			tags.pop();
		}
	}
};	//	class XmlStream

// Helper functions, they may be simply overwritten
// E.g. you may use _tstring instead of const TCHAR*

inline const XmlStream::Controller prolog() {
	return XmlStream::Controller(XmlStream::Controller::whatProlog);
}

inline const XmlStream::Controller tag() {
	return XmlStream::Controller(XmlStream::Controller::whatTag);
}

inline const XmlStream::Controller tag(const TCHAR* const tag_name) {
	return XmlStream::Controller(XmlStream::Controller::whatTag, tag_name);
}

inline const XmlStream::Controller endtag() {
	return XmlStream::Controller(XmlStream::Controller::whatTagEnd);
}

inline const XmlStream::Controller endtag(const TCHAR* const tag_name) {
	return XmlStream::Controller(XmlStream::Controller::whatTagEnd, tag_name);
}

inline const XmlStream::Controller attr(const TCHAR* const attr_name) {
	return XmlStream::Controller(XmlStream::Controller::whatAttribute, attr_name);
}

inline const XmlStream::Controller chardata() {
	return XmlStream::Controller(XmlStream::Controller::whatCharData);
}

}	// namespace

#endif  //  __XMLWRITER_H_2CC1D410_9DE8_4452_B8F0_F987EF152E06
