/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009-2012  Zotero
	                         Center for History and New Media
						     George Mason University, Fairfax, Virginia, USA
						     http://zotero.org
	
	Zotero is free software: you can redistribute it and/or modify
	it under the terms of the GNU Affero General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	Zotero is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU Affero General Public License for more details.
	
	You should have received a copy of the GNU Affero General Public License
	along with Zotero.  If not, see <http://www.gnu.org/licenses/>.
    
    Permission is granted to link statically the libraries included with
    a stock copy of Microsoft Windows. This library may not be linked, 
    directly or indirectly, with any other proprietary code.
    
    ***** END LICENSE BLOCK *****
*/

#include "zoteroWinWordIntegration.h"

#include <vector>

static COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
static wchar_t* FIELD_PREFIXES[] = {L" ADDIN ZOTERO_", L" CSL_", NULL};
static wchar_t* BOOKMARK_PREFIXES[] = {L"ZOTERO_", L"CSL_", NULL};
static const long WD_STYLE_TYPE_CHARACTER = 2;
static const long WD_STYLE_DEFAULT_PARAGRAPH_FONT = -66;
static const short WD_FIELD_TOA_ENTRY = 74;
static const long WD_COLLAPSE_END = 0;

statusCode isWholeNote(field_t* field, bool* returnValue);
statusCode setTextAndNoteLocations(field_t* field);

static bool getCapsStyleSmallCaps(document_t *doc, const wchar_t smallCapsStyleName[], bool *returnValue) {
	if(!smallCapsStyleName || !smallCapsStyleName[0]) return false;
	try {
		CStyles styles = doc->comDoc.get_Styles();
		CStyle style(styles.Item(smallCapsStyleName));
		if(style.get_Type() != WD_STYLE_TYPE_CHARACTER) return false;
		CFont0 styleFont = style.get_Font();
		*returnValue = styleFont.get_SmallCaps() != 0;
		return true;
	}
	catch(CException* e) {
		e->Delete();
		return false;
	}
	catch(...) {
		return false;
	}
}

static void applyCapsStyleToRange(CRange *baseRange, long start, long end,
								  const wchar_t smallCapsStyleName[],
								  bool capsStyleUsesSmallCaps) {
	if(start >= end) return;
	CRange runRange = baseRange->get_Duplicate();
	runRange.SetRange(start, end);
	runRange.put_Style(smallCapsStyleName);
	if(!capsStyleUsesSmallCaps) {
		CFont0 runFont = runRange.get_Font();
		runFont.put_SmallCaps(0);
	}
}

static void applyCapsStyleToSmallCaps(document_t *doc, CRange *range,
									  const wchar_t smallCapsStyleName[]) {
	bool capsStyleUsesSmallCaps = false;
	if(!getCapsStyleSmallCaps(doc, smallCapsStyleName, &capsStyleUsesSmallCaps)) return;

	long start = range->get_Start();
	long end = range->get_End();
	long runStart = -1;

	for(long pos = start; pos < end; pos++) {
		CRange charRange = range->get_Duplicate();
		charRange.SetRange(pos, pos + 1);
		CFont0 charFont = charRange.get_Font();
		bool isSmallCaps = charFont.get_SmallCaps() != 0;

		if(isSmallCaps && runStart == -1) {
			runStart = pos;
		}
		else if(!isSmallCaps && runStart != -1) {
			applyCapsStyleToRange(range, runStart, pos, smallCapsStyleName, capsStyleUsesSmallCaps);
			runStart = -1;
		}
	}

	if(runStart != -1) {
		applyCapsStyleToRange(range, runStart, end, smallCapsStyleName, capsStyleUsesSmallCaps);
	}
}

static void insertRTFIntoRange(document_t *doc, CRange *range, const wchar_t string[],
							   const wchar_t smallCapsStyleName[] = L"") {
	char* utf8String;
	int nBytes = WideCharToMultiByte(CP_UTF8, 0, string, -1, NULL, 0, NULL, NULL);
	utf8String = new char[nBytes];
	WideCharToMultiByte(CP_UTF8, 0, string, -1, utf8String, nBytes, NULL, NULL);

	DWORD nWritten;
	HANDLE tempFileHandle = getTemporaryFile();
	WriteFile(tempFileHandle, utf8String, nBytes-1, &nWritten, NULL);
	SetEndOfFile(tempFileHandle);
	delete[] utf8String;

	range->put_Text(L"");
	insertTemporaryFile(range);

	if(!wcsstr(string, L"\\\r") && !wcsstr(string, L"\\par") && !wcsstr(string, L"\\\n")) {
		CRange toDelete = range->get_Duplicate();
		toDelete.Collapse(0);
		toDelete.MoveStart(1, -1);
		if(toDelete.get_Text() != L"\x0d") {
			toDelete.Collapse(0);
			toDelete.MoveEnd(1, 1);
		}
		toDelete.put_Text(L"");
	}

	if(wcsstr(string, L"\\scaps")) {
		applyCapsStyleToSmallCaps(doc, range, smallCapsStyleName);
	}
}

static void clearTOAMarks(CRange *range) {
	CFields fields = range->get_Fields();
	long count = fields.get_Count();
	for(long i = count; i >= 1; i--) {
		CField field = fields.Item(i);
		if(field.get_Type() == WD_FIELD_TOA_ENTRY) {
			field.Delete();
		}
	}
}

static void removeTextOnce(CString *text, const CString& toRemove) {
	if(toRemove.IsEmpty()) return;
	int pos = text->Find(toRemove);
	if(pos != -1) {
		text->Delete(pos, toRemove.GetLength());
	}
}

static void removeRangeTextOnce(CString *text, CRange range) {
	removeTextOnce(text, range.get_Text());
}

static void removeTOAFieldTextOnce(CString *text, CField *field) {
	try {
		removeRangeTextOnce(text, field->get_Code());
	}
	catch(CException* e) {
		e->Delete();
	}
	try {
		removeRangeTextOnce(text, field->get_Result());
	}
	catch(CException* e) {
		e->Delete();
	}
}

static void removeFieldControlChars(CString *text) {
	for(int i = text->GetLength() - 1; i >= 0; i--) {
		wchar_t ch = text->GetAt(i);
		if(ch == L'\x13' || ch == L'\x14' || ch == L'\x15') {
			text->Delete(i, 1);
		}
	}
}

static CString getTextWithoutTOAMarks(CRange *range) {
	CString text = range->get_Text();
	CFields fields = range->get_Fields();
	long count = fields.get_Count();
	bool removedTOAMark = false;
	for(long i = count; i >= 1; i--) {
		CField nestedField = fields.Item(i);
		if(nestedField.get_Type() != WD_FIELD_TOA_ENTRY) {
			continue;
		}
		removedTOAMark = true;
		removeTOAFieldTextOnce(&text, &nestedField);
	}
	if(removedTOAMark) {
		removeFieldControlChars(&text);
	}
	return text;
}

struct TOAFormatState {
	bool bold;
	bool italic;
	bool underline;
	bool smallCaps;
	bool superscript;
	bool subscript;

	TOAFormatState() :
		bold(false),
		italic(false),
		underline(false),
		smallCaps(false),
		superscript(false),
		subscript(false) {}
};

struct TOAFormatRun {
	long start;
	long end;
	TOAFormatState state;
};

struct TOAFormattedText {
	CString text;
	std::vector<TOAFormatRun> runs;
};

static bool sameTOAFormatState(const TOAFormatState& a, const TOAFormatState& b) {
	return a.bold == b.bold
		&& a.italic == b.italic
		&& a.underline == b.underline
		&& a.smallCaps == b.smallCaps
		&& a.superscript == b.superscript
		&& a.subscript == b.subscript;
}

static bool hasTOAFormat(const TOAFormatState& state) {
	return state.bold || state.italic || state.underline || state.smallCaps
		|| state.superscript || state.subscript;
}

static void appendTOAChar(TOAFormattedText *formatted, wchar_t ch, const TOAFormatState& state) {
	long pos = formatted->text.GetLength();
	formatted->text.AppendChar(ch);
	if(!hasTOAFormat(state)) return;

	if(!formatted->runs.empty()) {
		TOAFormatRun& lastRun = formatted->runs.back();
		if(lastRun.end == pos && sameTOAFormatState(lastRun.state, state)) {
			lastRun.end++;
			return;
		}
	}

	TOAFormatRun run;
	run.start = pos;
	run.end = pos + 1;
	run.state = state;
	formatted->runs.push_back(run);
}

static void appendTOALineBreak(TOAFormattedText *formatted, const TOAFormatState& state) {
	long length = formatted->text.GetLength();
	if(length == 0 || formatted->text[length - 1] == L'\v') return;
	appendTOAChar(formatted, L'\v', state);
}

static void trimTOATrailingLineBreaks(TOAFormattedText *formatted) {
	while(!formatted->text.IsEmpty()) {
		long length = formatted->text.GetLength();
		wchar_t ch = formatted->text[length - 1];
		if(ch != L'\v' && ch != L' ' && ch != L'\t') break;
		formatted->text.Truncate(length - 1);
	}

	while(!formatted->runs.empty() && formatted->runs.back().start >= formatted->text.GetLength()) {
		formatted->runs.pop_back();
	}
	if(!formatted->runs.empty() && formatted->runs.back().end > formatted->text.GetLength()) {
		formatted->runs.back().end = formatted->text.GetLength();
	}
}

static TOAFormattedText parseTOARTF(const wchar_t string[]) {
	TOAFormattedText formatted;
	TOAFormatState state;
	std::vector<TOAFormatState> stack;
	long length = wcslen(string);

	for(long i = 0; i < length; i++) {
		wchar_t ch = string[i];
		if(ch == L'{') {
			stack.push_back(state);
			continue;
		}
		if(ch == L'}') {
			if(!stack.empty()) {
				state = stack.back();
				stack.pop_back();
			}
			continue;
		}
		if(ch == L'\r' || ch == L'\n') {
			continue;
		}
		if(ch != L'\\') {
			appendTOAChar(&formatted, ch, state);
			continue;
		}

		if(i + 1 >= length) break;
		wchar_t next = string[++i];
		if(next == L'\\' || next == L'{' || next == L'}') {
			appendTOAChar(&formatted, next, state);
			continue;
		}

		long controlStart = i;
		while(i < length && ((string[i] >= L'a' && string[i] <= L'z') || (string[i] >= L'A' && string[i] <= L'Z'))) {
			i++;
		}
		CString control;
		control.SetString(string + controlStart, i - controlStart);

		bool negative = false;
		if(i < length && string[i] == L'-') {
			negative = true;
			i++;
		}

		bool hasNumber = false;
		long number = 0;
		while(i < length && string[i] >= L'0' && string[i] <= L'9') {
			hasNumber = true;
			number = number * 10 + (string[i] - L'0');
			i++;
		}
		if(negative) number = -number;

		if(i < length && (string[i] == L' ' || string[i] == L'\r' || string[i] == L'\n')) {
			// Control-word delimiter, not literal text.
		}
		else {
			i--;
		}

		if(control == L"rtf" || control == L"ansi" || control == L"deff"
				|| control == L"uc" || control == L"pard" || control == L"plain") {
			continue;
		}
		if(control == L"i") {
			state.italic = !hasNumber || number != 0;
			continue;
		}
		if(control == L"b") {
			state.bold = !hasNumber || number != 0;
			continue;
		}
		if(control == L"ul") {
			state.underline = !hasNumber || number != 0;
			continue;
		}
		if(control == L"ulnone") {
			state.underline = false;
			continue;
		}
		if(control == L"scaps") {
			state.smallCaps = !hasNumber || number != 0;
			continue;
		}
		if(control == L"super") {
			state.superscript = true;
			state.subscript = false;
			continue;
		}
		if(control == L"sub") {
			state.subscript = true;
			state.superscript = false;
			continue;
		}
		if(control == L"nosupersub") {
			state.superscript = false;
			state.subscript = false;
			continue;
		}
		if(control == L"tab") {
			appendTOAChar(&formatted, L'\t', state);
			continue;
		}
		if(control == L"line" || control == L"par") {
			appendTOALineBreak(&formatted, state);
			continue;
		}
		if(control == L"u" && hasNumber) {
			appendTOAChar(&formatted, (wchar_t) number, state);
			if(i + 2 < length && string[i + 1] == L'{' && string[i + 2] == L'}') {
				i += 2;
			}
			continue;
		}
	}

	trimTOATrailingLineBreaks(&formatted);
	return formatted;
}

static void appendFieldArgument(const CString& text, CString *fieldCode, std::vector<long> *offsets) {
	for(long i = 0; i < text.GetLength(); i++) {
		wchar_t ch = text[i];
		offsets->push_back(fieldCode->GetLength());
		if(ch == L'\\' || ch == L'"') {
			fieldCode->AppendChar(L'\\');
		}
		fieldCode->AppendChar(ch);
	}
}

static void applyCapsFontDirectly(document_t *doc, CRange *range, const wchar_t smallCapsStyleName[]) {
	if(!smallCapsStyleName || !smallCapsStyleName[0]) {
		CFont0 font = range->get_Font();
		font.put_SmallCaps(1);
		return;
	}
	try {
		CStyles styles = doc->comDoc.get_Styles();
		CStyle capsStyle(styles.Item(smallCapsStyleName));
		if(capsStyle.get_Type() != WD_STYLE_TYPE_CHARACTER) {
			CFont0 font = range->get_Font();
			font.put_SmallCaps(1);
			return;
		}

		CStyle defaultStyle(styles.Item(WD_STYLE_DEFAULT_PARAGRAPH_FONT));
		CFont0 capsFont = capsStyle.get_Font();
		CFont0 defaultFont = defaultStyle.get_Font();
		CFont0 rangeFont = range->get_Font();

		CString name = capsFont.get_Name();
		if(!name.IsEmpty() && name != defaultFont.get_Name()) rangeFont.put_Name(name);
		CString nameAscii = capsFont.get_NameAscii();
		if(!nameAscii.IsEmpty() && nameAscii != defaultFont.get_NameAscii()) rangeFont.put_NameAscii(nameAscii);
		CString nameOther = capsFont.get_NameOther();
		if(!nameOther.IsEmpty() && nameOther != defaultFont.get_NameOther()) rangeFont.put_NameOther(nameOther);
		CString nameFarEast = capsFont.get_NameFarEast();
		if(!nameFarEast.IsEmpty() && nameFarEast != defaultFont.get_NameFarEast()) rangeFont.put_NameFarEast(nameFarEast);

		float size = capsFont.get_Size();
		if(size > 0 && size != defaultFont.get_Size()) rangeFont.put_Size(size);
		float spacing = capsFont.get_Spacing();
		if(spacing != defaultFont.get_Spacing()) rangeFont.put_Spacing(spacing);
		long scaling = capsFont.get_Scaling();
		if(scaling > 0 && scaling != defaultFont.get_Scaling()) rangeFont.put_Scaling(scaling);
		long position = capsFont.get_Position();
		if(position != defaultFont.get_Position()) rangeFont.put_Position(position);
		float kerning = capsFont.get_Kerning();
		if(kerning != defaultFont.get_Kerning()) rangeFont.put_Kerning(kerning);

		if(capsFont.get_Bold() != defaultFont.get_Bold()) rangeFont.put_Bold(capsFont.get_Bold());
		if(capsFont.get_Italic() != defaultFont.get_Italic()) rangeFont.put_Italic(capsFont.get_Italic());
		if(capsFont.get_SmallCaps() != defaultFont.get_SmallCaps()) rangeFont.put_SmallCaps(capsFont.get_SmallCaps());
		if(capsFont.get_AllCaps() != defaultFont.get_AllCaps()) rangeFont.put_AllCaps(capsFont.get_AllCaps());
	}
	catch(CException* e) {
		e->Delete();
		CFont0 font = range->get_Font();
		font.put_SmallCaps(1);
	}
	catch(...) {
		CFont0 font = range->get_Font();
		font.put_SmallCaps(1);
	}
}

static void applyTOAFormatting(document_t *doc, CField *toaField, const CString& fieldCode,
							   const TOAFormattedText& formatted,
							   const std::vector<long>& offsets,
							   const wchar_t smallCapsStyleName[]) {
	if(offsets.empty()) return;

	CRange codeRange = toaField->get_Code();
	CString codeText = codeRange.get_Text();
	long base = codeText.Find(fieldCode);
	if(base == -1) return;

	long codeStart = codeRange.get_Start() + base;
	for(size_t i = 0; i < formatted.runs.size(); i++) {
		TOAFormatRun run = formatted.runs[i];
		if(run.start < 0 || run.end <= run.start || run.end > (long) offsets.size()) continue;

		CRange runRange = codeRange.get_Duplicate();
		runRange.SetRange(codeStart + offsets[run.start], codeStart + offsets[run.end - 1] + 1);
		CFont0 font = runRange.get_Font();
		if(run.state.bold) font.put_Bold(1);
		if(run.state.italic) font.put_Italic(1);
		if(run.state.underline) runRange.put_Underline(1);
		if(run.state.superscript) font.put_Superscript(1);
		if(run.state.subscript) font.put_Subscript(1);
		if(run.state.smallCaps) applyCapsFontDirectly(doc, &runRange, smallCapsStyleName);
	}
}

static void insertTOAField(document_t *doc, CRange *range, const wchar_t shortCitation[],
						   const wchar_t longCitation[], unsigned short category,
						   bool isInitial, const wchar_t smallCapsStyleName[]) {
	TOAFormattedText shortText = parseTOARTF(shortCitation ? shortCitation : L"");
	TOAFormattedText longText = parseTOARTF(longCitation ? longCitation : L"");

	CString fieldCode;
	std::vector<long> shortOffsets;
	std::vector<long> longOffsets;

	if(isInitial && !longText.text.IsEmpty()) {
		fieldCode.Append(L"\\l \"");
		appendFieldArgument(longText.text, &fieldCode, &longOffsets);
		fieldCode.Append(L"\" ");
	}
	fieldCode.Append(L"\\s \"");
	appendFieldArgument(shortText.text, &fieldCode, &shortOffsets);
	fieldCode.Append(L"\" ");
	CString categoryCode;
	categoryCode.Format(L"\\c %d", category);
	fieldCode.Append(categoryCode);

	CFields fields = range->get_Fields();
	CField toaField = fields.Add(*range, COleVariant((short) WD_FIELD_TOA_ENTRY),
		COleVariant(fieldCode), COleVariant((short) VARIANT_FALSE, VT_BOOL));

	applyTOAFormatting(doc, &toaField, fieldCode, shortText, shortOffsets, smallCapsStyleName);
	applyTOAFormatting(doc, &toaField, fieldCode, longText, longOffsets, smallCapsStyleName);
}

static void insertTOAMark(field_t* field, const wchar_t shortCitation[],
						  const wchar_t longCitation[], unsigned short category,
						  bool isInitial, const wchar_t smallCapsStyleName[]) {
	CRange insertRange = field->comContentRange.get_Duplicate();
	insertRange.Collapse(WD_COLLAPSE_END);

	insertTOAField(field->doc, &insertRange, shortCitation,
		longCitation, category, isInitial, smallCapsStyleName);
}

// Allocates a field structure based on a CField, optionally checking to make
// sure that the field code actually matches a Zotero field.
statusCode initField(document_t *doc, CField comField, short noteType,
					 bool ignoreCode, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CRange comCodeRange = comField.get_Code();
	
	field_t *field = NULL;
	if(ignoreCode) {
		field = new field_t;
		field->code = NULL;
	} else {
		// Read code
		ENSURE_OK(prepareReadFieldCode(doc));
		CString rawCode = comCodeRange.get_Text();
		
		if(rawCode) {
			// Check that this field is a valid Zotero field
			for(unsigned short i=0; FIELD_PREFIXES[i] != NULL; i++) {
				int location = rawCode.Find(FIELD_PREFIXES[i]);
				if(location != -1) {
					field = new field_t;
					
					// If field code is all caps, make sure text object isn't in
					// all caps mode
					CString upperCaseCode = rawCode;
					upperCaseCode.MakeUpper();
					if(rawCode == upperCaseCode) {
						try {
							CFont0 comFont = comCodeRange.get_Font();
							if(comFont.get_AllCaps()) {
								comFont.put_AllCaps(false);
								rawCode = comCodeRange.get_Text();
							}
						} catch(CException* e) {
							e->Delete();
						}
					}
					
					// Get code
					int rawCodeLength = rawCode.GetLength();
					int codeStart = location + ((int) wcslen(FIELD_PREFIXES[i]));
					int codeLength = rawCodeLength - codeStart;
					
					// Sometimes there is a space at the end of the code, which
					// we ignore here
					if(rawCode.GetAt(rawCodeLength-1) == L' ') {
						codeLength--;
					}

					CString code = rawCode.Mid(codeStart, codeLength);
					field->code = _wcsdup(code);
					
					break;
				}
			}
		}
	}
	
	if(field) {
		field->text = NULL;
		field->bookmarkName = NULL;
		field->comBookmark = NULL;
		field->adjacent = false;
		
		field->doc = doc;
		field->comField = comField;
		field->comCodeRange = comCodeRange;
		field->comContentRange = comField.get_Result();
		field->noteType = noteType;
		setTextAndNoteLocations(field);
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Allocates a field structure based on a WordBookmark
statusCode initBookmark(document_t *doc, CBookmark0 comBookmark, short noteType,
						bool ignoreCode, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	field_t *field = NULL;
	
	// Read code
	CString bookmarkName = comBookmark.get_Name();
	
	if(ignoreCode) {
		field = new field_t;
		field->code = NULL;
	} else {
		// Check that this field is a valid Zotero field
		for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
			int location = bookmarkName.Find(BOOKMARK_PREFIXES[i]);
			if(location != -1) {
				// Get code from preferences
				CString propertyValue;
				ENSURE_OK(getProperty(doc, bookmarkName, &propertyValue));
				
				// Check that preferences code is valid
				for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
					int location = propertyValue.Find(BOOKMARK_PREFIXES[i]);
					if(location != -1) {
						field = new field_t;

						// Get code
						int rawCodeLength = propertyValue.GetLength();
						int codeStart = location + ((int) wcslen(BOOKMARK_PREFIXES[i]));
						int codeLength = rawCodeLength - codeStart;

						CString code = propertyValue.Mid(codeStart, codeLength);
						field->code = _wcsdup(code);
						break;
					}
				}
				
				if(field) break;
			}
		}
		
	}
	
	if(field) {
		field->text = NULL;
		field->comField = NULL;
		field->adjacent = false;
		
		field->bookmarkName = _wcsdup(bookmarkName);
		field->doc = doc;
		field->comBookmark = comBookmark;
		
		// Get the field contents
		field->comContentRange = comBookmark.get_Range();
		field->comCodeRange = field->comContentRange;
		field->noteType = noteType;
		setTextAndNoteLocations(field);
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Frees a field struct. This does not free the corresponding field list node.
void __stdcall freeField(field_t* field) {
	if(field->text) delete field->text;
	if(field->code) delete field->code;
	if(field->bookmarkName) free(field->bookmarkName);
	delete field;
}

// Deletes this field from the document
statusCode __stdcall deleteField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	bool wholeNote;
	ENSURE_OK(isWholeNote(field, &wholeNote));
	CSelection selection = field->doc->comWindow.get_Selection();
	CRange selectionRange = selection.get_Range();
	if(wholeNote) {
		// If the note contains only this field, delete the note
		if(field->noteType == NOTE_FOOTNOTE) {
			CRange noteRange = field->comFootnote.get_Range();
			// If we remove a note and the selection cursor is in it, the document ends up without a valid selection
			// so we put the selection right after the note in the reference.
			if (selection.get_Start() >= noteRange.get_Start() && selection.get_End() <= noteRange.get_End()) {
				CRange refRange = field->comFootnote.get_Reference();
				CRange dupRange = refRange.get_Duplicate();
				dupRange.Collapse(0);
				dupRange.Select();
				field->doc->insertTextIntoNote = field->noteType;
			}
			field->comFootnote.Delete();
		} else if(field->noteType == NOTE_ENDNOTE) {
			CRange noteRange = field->comEndnote.get_Range();
			if (selection.get_Start() >= noteRange.get_Start() && selection.get_End() <= noteRange.get_End()) {
				CRange refRange = field->comEndnote.get_Reference();
				CRange dupRange = refRange.get_Duplicate();
				dupRange.Collapse(0);
				dupRange.Select();
				field->doc->insertTextIntoNote = field->noteType;
			}
			field->comEndnote.Delete();
		}
	} else if(field->comBookmark) {
		field->comContentRange.put_Text(L"");
	} else {
		field->comField.Delete();
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Selects this field
statusCode __stdcall selectField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(field->doc, true);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Removes a field code
statusCode __stdcall removeCode(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		field->comBookmark.Delete();
		setProperty(field->doc, field->bookmarkName, L"");
	} else {
		field->comField.Unlink();
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Gets text inside this field. DO NOT FREE THE RETURN VALUE!
statusCode __stdcall getText(field_t* field, wchar_t** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	CString fieldText = getTextWithoutTOAMarks(&field->comContentRange);
	*returnValue = field->text = _wcsdup(fieldText);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets text of this field
statusCode __stdcall setText(field_t* field, const wchar_t string[], bool isRich,
							 const wchar_t smallCapsStyleName[]) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(field->doc, false);
	
	// Get current font size and name
	CFont0 comFont = field->comContentRange.get_Font();
	float fontSize = comFont.get_Size();
	CString fontName = comFont.get_Name();

	// If the font name from the range is empty (happens with mixed fonts,
	// e.g. when inserting text with mixed latin and chinese characters),
	// fall back to the font name from the current style. Otherwise Word
	// defaults to Times New Roman when we put an empty font name back on.
	if (fontName.IsEmpty()) {
		CComVariant varStyle = field->comContentRange.get_Style();
		if (varStyle.vt == VT_DISPATCH && varStyle.pdispVal != NULL) {
			CStyle comStyle(varStyle.pdispVal);
			CFont0 styleFont = comStyle.get_Font();
			fontName = styleFont.get_Name();
		}
	}

	// Check if we need to restore cursor position after insert (bookmarks only)
	bool restoreSelectionToEnd = false;
	if(field->comBookmark) {
		CSelection comSelection = (field->doc)->comWindow.get_Selection();
		CRange comSelectionRange = comSelection.get_Range();
		CRange comTestRange = field->comContentRange.get_Duplicate();
		comTestRange.Collapse(0 /*wdCollapseEnd*/);
		if(comSelectionRange.IsEqual(comTestRange)) {
			restoreSelectionToEnd = true;
		}
	}

	if(isRich) {
		// Get a temp file

		char* utf8String;
		int nBytes;
		if(field->comBookmark && wcslen(string) > 6) {
			// InsertFile method will clobber the bookmark, so add it to the RTF
			CString insertString;
			insertString.Format(L"{\\rtf {\\bkmkstart %s}{%s{\\bkmkend %s}}", field->bookmarkName, string+6, field->bookmarkName);
			nBytes = WideCharToMultiByte(CP_UTF8, 0, insertString, -1, NULL, 0, NULL, NULL);
			utf8String = new char[nBytes];
			WideCharToMultiByte(CP_UTF8, 0, insertString, -1, utf8String, nBytes, NULL, NULL);
		} else {
			// Convert to UTF-8
			nBytes = WideCharToMultiByte(CP_UTF8, 0, string, -1, NULL, 0, NULL, NULL);
			utf8String = new char[nBytes];
			WideCharToMultiByte(CP_UTF8, 0, string, -1, utf8String, nBytes, NULL, NULL);
		}

		// Open and write file
		DWORD nWritten;
		HANDLE tempFileHandle = getTemporaryFile();
		WriteFile(tempFileHandle, utf8String, nBytes-1, &nWritten, NULL);
		SetEndOfFile(tempFileHandle);
		delete[] utf8String;

		// Read from file into range
		if(!(field->comBookmark) && (field->doc)->wordVersion >= 15) {
			// In Word 2013, text does not get inserted into ranges
			field->comContentRange.put_Text(L"  ");
			CRange toDelete = field->comContentRange.get_Duplicate();
			toDelete.Collapse(0);
			CRange comDupRange = field->comContentRange.get_Duplicate();
			comDupRange.MoveEnd(1, -1);
			insertTemporaryFile(&comDupRange);
			toDelete.MoveStart(1, -1);
			toDelete.put_Text(L"");
		} else {
			field->comContentRange.put_Text(L"");
			insertTemporaryFile(&field->comContentRange);

			if(field->comBookmark) {
				field->comContentRange = field->comBookmark.get_Range();
				field->comCodeRange = field->comContentRange;
			} else {
				field->comContentRange = field->comField.get_Result();
			}
		}

		// We need to reset the font here, otherwise in mixed-font situations (e.g. when citing
		// items that contain english and chinese characters) the font somehow goes to Times New Roman
		CFont0 comFont = field->comContentRange.get_Font();
		comFont.put_Name(fontName);
		comFont.put_Size(fontSize);

		// Need to delete the return that gets added at the end, but only if there are no
		// returns within the text to be inserted
		if(!wcsstr(string, L"\\\r") && !wcsstr(string, L"\\par") && !wcsstr(string, L"\\\n")) {
			CRange toDelete = field->comContentRange.get_Duplicate();
			toDelete.Collapse(0);
			toDelete.MoveStart(1, -1);
			if(toDelete.get_Text() != L"\x0d") {
				toDelete.Collapse(0);
				toDelete.MoveEnd(1, 1);
			}
			toDelete.put_Text(L"");
		}

		if(wcsncmp(field->code, L"BIBL", 4) == 0) {
			setStyle(field->doc, &field->comContentRange, BIBLIOGRAPHY_STYLE_ENUM, BIBLIOGRAPHY_STYLE_NAME);
		}

		if(wcsstr(string, L"\\scaps")) {
			applyCapsStyleToSmallCaps(field->doc, &field->comContentRange, smallCapsStyleName);
		}
	} else {
		CFont0 comFont = field->comContentRange.get_Font();
		comFont.Reset();
		field->comContentRange.put_Text(string);
		if(field->comBookmark) {
			// Setting the text of the bookmark erases it, so we need to recreate it
			CBookmarks comBookmarks = field->doc->comDoc.get_Bookmarks();
			field->comBookmark = comBookmarks.Add(field->bookmarkName, field->comContentRange);
			field->comContentRange = field->comBookmark.get_Range();
		}
		// Put font back on
		comFont.put_Name(fontName);
		comFont.put_Size(fontSize);
	}

	// Restore the selection to the end of a bookmark
	if(restoreSelectionToEnd) {
		CRange comRange = field->comContentRange.get_Duplicate();
		comRange.Collapse(0 /*wdCollapseEnd*/);
		comRange.Select();
	}

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets Table of Authorities marks inside this field result.
statusCode __stdcall setTOAMarks(field_t* field, const wchar_t* shortCitations[],
								 const wchar_t* longCitations[], unsigned short categories[],
								 bool isInitial[], unsigned long count,
								 const wchar_t smallCapsStyleName[]) {
	HANDLE_EXCEPTIONS_BEGIN
	setScreenUpdatingStatus(field->doc, false);

	clearTOAMarks(&field->comContentRange);
	for(unsigned long i = 0; i < count; i++) {
		insertTOAMark(field, shortCitations[i], longCitations[i], categories[i],
			isInitial[i], smallCapsStyleName);
	}

	if(field->comBookmark) {
		field->comContentRange = field->comBookmark.get_Range();
		field->comCodeRange = field->comContentRange;
	}
	else {
		field->comContentRange = field->comField.get_Result();
		field->comCodeRange = field->comField.get_Code();
	}
	setTextAndNoteLocations(field);

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets the field code
statusCode __stdcall setCode(field_t *field, const wchar_t code[]) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		CString rawCode;
		rawCode.Format(L"%s%s", BOOKMARK_PREFIXES[0], code);
		ENSURE_OK(setProperty(field->doc, field->bookmarkName, rawCode));
	} else {
		CString rawCode;
		rawCode.Format(L"%s%s ", FIELD_PREFIXES[0], code);
		field->comCodeRange.put_Text(rawCode);
	}
	
	// Store code in struct
	if(field->code) free(field->code);
	field->code = _wcsdup(code);
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns the index of the note in which this field resides
statusCode __stdcall getNoteIndex(field_t* field, unsigned long *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == NOTE_FOOTNOTE) {
		*returnValue = field->comFootnote.get_Index();
	} else if(field->noteType == NOTE_ENDNOTE){ 
		*returnValue = field->comEndnote.get_Index();
	} else {
		*returnValue = 0;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns whether the field is adjacent to the next field
statusCode __stdcall isAdjacentToNextField(field_t* field, bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = field->adjacent;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Compares two fields to determine which comes before which
statusCode compareFields(field_t* a, field_t* b, short *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(a->textLocation < b->textLocation) {
		*returnValue = -1;
		return STATUS_OK;
	} else if(b->textLocation < a->textLocation) {
		*returnValue = 1;
		return STATUS_OK;
	}
	
	// Compare positions inside a footnote
	if(a->noteType && b->noteType) {
		if(a->noteLocation < b->noteLocation) {
			*returnValue = -1;
			return STATUS_OK;
		} else if(b->noteLocation < a->noteLocation) {
			*returnValue = 1;
			return STATUS_OK;
		}
	}
	
	*returnValue = 0;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Converts this field to a different note type. The implementation should not expect to use the
// field structure for anything else after this happens.
statusCode convertToNoteType(field_t* field, short toNoteType) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == NOTE_FOOTNOTE && toNoteType == NOTE_ENDNOTE) {
		// Footnote to endnote
		CRange comRange = field->comFootnote.get_Range();
		CFootnotes comFootnotes = comRange.get_Footnotes();
		comFootnotes.Convert();
	} else if(field->noteType == NOTE_ENDNOTE && toNoteType == NOTE_FOOTNOTE) {
		// Endnote to footnote
		CRange comRange = field->comEndnote.get_Range();
		CEndnotes comEndnotes = comRange.get_Endnotes();
		comEndnotes.Convert();
	} else {
		CRange comRange;
		if(field->noteType && !toNoteType) {
			// Footnote or endnote to in-text
			bool wholeNote;
			ENSURE_OK(isWholeNote(field, &wholeNote));
			if(wholeNote) { // Replace reference with citation if this is the only one in the note
				CRange comRefRange, comNoteRange;
				if(field->noteType == NOTE_FOOTNOTE) {
					comRefRange = field->comFootnote.get_Reference();
					comNoteRange = field->comFootnote.get_Range();
				} else {
					comRefRange = field->comEndnote.get_Reference();
					comNoteRange = field->comEndnote.get_Range();
				}
				comRange = comRefRange.get_Duplicate();
				comRange.Collapse(0);
				comRange.put_FormattedText(comNoteRange);
				comRefRange.put_Text(L"");
			}
		} else if(!field->noteType && toNoteType) {
			// In-text to footnote or endnote
			// Get document
			comRange = field->comContentRange.get_Duplicate();
			comRange.Collapse(0);
			comRange.Move(1, 1);
		
			// Create a new note and get its range
			if(toNoteType == NOTE_FOOTNOTE) {
				CFootnotes notes = (field->doc)->comDoc.get_Footnotes();
				CFootnote note = notes.Add(comRange, covOptional, COleVariant(L""));
				comRange = note.get_Range();
			} else if(toNoteType == NOTE_ENDNOTE) {
				CEndnotes notes = (field->doc)->comDoc.get_Endnotes();
				CEndnote note = notes.Add(comRange, covOptional, COleVariant(L""));
				comRange = note.get_Range();
			}
		
			// Put formatted text in the range
			CRange fieldRange;
			ENSURE_OK(getFieldRange(field, &fieldRange));
			comRange.put_FormattedText(fieldRange);
			deleteField(field);
		}

		// If a bookmark, re-create the mark
		if(field->bookmarkName) {
			CBookmarks comBookmarks = field->doc->comDoc.get_Bookmarks();
			field->comBookmark = comBookmarks.Add(field->bookmarkName, comRange);
		}
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode isWholeNote(field_t* field, bool* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType) {
		CRange noteRange;

		if(field->noteType == NOTE_FOOTNOTE) {
			noteRange = field->comFootnote.get_Range();
		} else if(field->noteType == NOTE_ENDNOTE) {
			noteRange = field->comEndnote.get_Range();
		}
		
		CRange testRange;
		ENSURE_OK(getFieldRange(field, &testRange));
		*returnValue = noteRange.IsEqual(testRange) != 0;
	} else {
		*returnValue = false;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Get a range encompassing both the code range and the content range
statusCode getFieldRange(field_t* field, CRange* testRange) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->comBookmark) {
		*testRange = field->comContentRange.get_Duplicate();
	} else {
		*testRange = field->comCodeRange.get_Duplicate();
		testRange->MoveStart(1, -1);
		testRange->put_End(field->comContentRange.get_End()+1);
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets noteType, textLocation, noteLocation, comFootnote, and comEndnote
statusCode setTextAndNoteLocations(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	if(field->noteType == -1) {
		long storyType = field->comCodeRange.get_StoryType();
		if(storyType == 2) {
			field->noteType = NOTE_FOOTNOTE;
		} else if(storyType == 3) {
			field->noteType = NOTE_ENDNOTE;
		} else {
			field->noteType = 0;
		}
	}

	if(field->noteType == NOTE_FOOTNOTE) {
		CFootnotes comFootnotes = field->comContentRange.get_Footnotes();
		field->comFootnote = comFootnotes.Item(1);
		field->comEndnote = NULL;
		CRange comNoteReference = field->comFootnote.get_Reference();
		field->textLocation = comNoteReference.get_Start();
		field->noteLocation = field->comCodeRange.get_Start();
	} else if(field->noteType == NOTE_ENDNOTE){ 
		CFootnotes comEndnotes = field->comContentRange.get_Endnotes();
		field->comEndnote = comEndnotes.Item(1);
		field->comFootnote = NULL;
		CRange comNoteReference = field->comEndnote.get_Reference();
		field->textLocation = comNoteReference.get_Start();
		field->noteLocation = field->comCodeRange.get_Start();
	} else {
		field->comFootnote = NULL;
		field->comEndnote = NULL;
		field->textLocation = field->comCodeRange.get_Start();
		field->noteLocation = field->textLocation;
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}
