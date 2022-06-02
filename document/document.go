package document

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"image"
	"image/jpeg"
	"io"
	"math/rand"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"unicode"

	"github.com/unidoc/unioffice"
	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/common"
	"github.com/unidoc/unioffice/common/axcontrol"
	"github.com/unidoc/unioffice/common/logger"
	"github.com/unidoc/unioffice/common/tempstorage"
	"github.com/unidoc/unioffice/internal/formatutils"
	"github.com/unidoc/unioffice/internal/license"
	"github.com/unidoc/unioffice/measurement"
	"github.com/unidoc/unioffice/schema/schemas.microsoft.com/office/activeX"
	"github.com/unidoc/unioffice/schema/soo/dml"
	dmlChart "github.com/unidoc/unioffice/schema/soo/dml/chart"
	"github.com/unidoc/unioffice/schema/soo/dml/picture"
	"github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes"
	"github.com/unidoc/unioffice/schema/soo/pkg/relationships"
	"github.com/unidoc/unioffice/schema/soo/wml"
	"github.com/unidoc/unioffice/schema/urn/schemas_microsoft_com/vml"
	"github.com/unidoc/unioffice/vmldrawing"
	"github.com/unidoc/unioffice/zippkg"
)

// VerticalAlign returns the value of run vertical align.
func (_effb RunProperties) VerticalAlignment() sharedTypes.ST_VerticalAlignRun {
	if _gdaac := _effb._gbdb.VertAlign; _gdaac != nil {
		return _gdaac.ValAttr
	}
	return 0
}

// SetFirstRow controls the conditional formatting for the first row in a table.
func (_deaff TableLook) SetFirstRow(on bool) {
	if !on {
		_deaff.ctTblLook.FirstRowAttr = &sharedTypes.ST_OnOff{}
		_deaff.ctTblLook.FirstRowAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	} else {
		_deaff.ctTblLook.FirstRowAttr = &sharedTypes.ST_OnOff{}
		_deaff.ctTblLook.FirstRowAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	}
}
func (_cfc *Document) insertNumberingFromStyleProperties(_gade Numbering, _gce ParagraphStyleProperties) {
	_accg := _gce.NumId()
	_ccgb := int64(-1)
	if _accg > -1 {
		for _, _fgfd := range _gade._cbag.Num {
			if _fgfd.NumIdAttr == _accg {
				if _fgfd.AbstractNumId != nil {
					_ccgb = _fgfd.AbstractNumId.ValAttr
					_cdcb := false
					for _, _bbfa := range _cfc.Numbering._cbag.Num {
						if _bbfa.NumIdAttr == _accg {
							_cdcb = true
							break
						}
					}
					if !_cdcb {
						_cfc.Numbering._cbag.Num = append(_cfc.Numbering._cbag.Num, _fgfd)
					}
					break
				}
			}
		}
		for _, _aggg := range _gade._cbag.AbstractNum {
			if _aggg.AbstractNumIdAttr == _ccgb {
				_eefa := false
				for _, _gdce := range _cfc.Numbering._cbag.AbstractNum {
					if _gdce.AbstractNumIdAttr == _ccgb {
						_eefa = true
						break
					}
				}
				if !_eefa {
					_cfc.Numbering._cbag.AbstractNum = append(_cfc.Numbering._cbag.AbstractNum, _aggg)
				}
				break
			}
		}
	}
}

// AddImage adds an image to the document package, returning a reference that
// can be used to add the image to a run and place it in the document contents.
func (_cffgg Footer) AddImage(i common.Image) (common.ImageRef, error) {
	var _cdec common.Relationships
	for _dgbg, _dgcb := range _cffgg._aegg._aba {
		if _dgcb == _cffgg._fcc {
			_cdec = _cffgg._aegg._fdf[_dgbg]
		}
	}
	_cfgea := common.MakeImageRef(i, &_cffgg._aegg.DocBase, _cdec)
	if i.Data == nil && i.Path == "" {
		return _cfgea, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0064\u0061t\u0061\u0020\u006f\u0072\u0020\u0061\u0020\u0070\u0061\u0074\u0068")
	}
	if i.Format == "" {
		return _cfgea, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0061\u0020v\u0061\u006c\u0069\u0064\u0020\u0066\u006f\u0072\u006d\u0061\u0074")
	}
	if i.Size.X == 0 || i.Size.Y == 0 {
		return _cfgea, errors.New("\u0069\u006d\u0061\u0067e\u0020\u006d\u0075\u0073\u0074\u0020\u0068\u0061\u0076\u0065 \u0061 \u0076\u0061\u006c\u0069\u0064\u0020\u0073i\u007a\u0065")
	}
	_cffgg._aegg.Images = append(_cffgg._aegg.Images, _cfgea)
	_bacd := fmt.Sprintf("\u006d\u0065d\u0069\u0061\u002fi\u006d\u0061\u0067\u0065\u0025\u0064\u002e\u0025\u0073", len(_cffgg._aegg.Images), i.Format)
	_cfdg := _cdec.AddRelationship(_bacd, unioffice.ImageType)
	_cfgea.SetRelID(_cfdg.X().IdAttr)
	return _cfgea, nil
}
func (_ebbf Endnote) content() []*wml.EG_ContentBlockContent {
	var _ecdb []*wml.EG_ContentBlockContent
	for _, _cgdgb := range _ebbf._fagg.EG_BlockLevelElts {
		_ecdb = append(_ecdb, _cgdgb.EG_ContentBlockContent...)
	}
	return _ecdb
}

// SetInsideHorizontal sets the interior horizontal borders to a specified type, color and thickness.
func (_fga CellBorders) SetInsideHorizontal(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_fga._gf.InsideH = wml.NewCT_Border()
	_feadc(_fga._gf.InsideH, t, c, thickness)
}

// AddHeader creates a header associated with the document, but doesn't add it
// to the document for display.
func (_abfe *Document) AddHeader() Header {
	_fbd := wml.NewHdr()
	_abfe._geb = append(_abfe._geb, _fbd)
	_gfdd := fmt.Sprintf("\u0068\u0065\u0061d\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", len(_abfe._geb))
	_abfe._dab.AddRelationship(_gfdd, unioffice.HeaderType)
	_abfe.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072\u0064\u002f"+_gfdd, "\u0061p\u0070l\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0076\u006e\u0064.\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c\u002e\u0068\u0065\u0061\u0064e\u0072\u002b\u0078\u006d\u006c")
	_abfe._cbfd = append(_abfe._cbfd, common.NewRelationships())
	return Header{_abfe, _fbd}
}

// Themes returns document's themes.
func (_dgafg *Document) Themes() []*dml.Theme { return _dgafg._ffbc }

type listItemInfo struct {
	FromStyle      *Style
	FromParagraph  *Paragraph
	AbstractNumId  *int64
	NumberingLevel *NumberingLevel
}

// RowProperties are the properties for a row within a table
type RowProperties struct{ _acgb *wml.CT_TrPr }

// FindNodeByText return node based on matched text and return a slice of node.
func (_aega *Nodes) FindNodeByText(text string) []Node {
	_gdcb := []Node{}
	for _, _bdbag := range _aega._gabfc {
		if strings.TrimSpace(_bdbag.Text()) == text {
			_gdcb = append(_gdcb, _bdbag)
		}
		_aaff := Nodes{_gabfc: _bdbag.Children}
		_gdcb = append(_gdcb, _aaff.FindNodeByText(text)...)
	}
	return _gdcb
}
func (_bba *Document) insertTable(_cgaa Paragraph, _gcc bool) Table {
	_efe := _bba.doc.Body
	if _efe == nil {
		return _bba.AddTable()
	}
	_cbd := _cgaa.X()
	for _cfa, _efc := range _efe.EG_BlockLevelElts {
		for _, _gdc := range _efc.EG_ContentBlockContent {
			for _ege, _gac := range _gdc.P {
				if _gac == _cbd {
					_ada := wml.NewCT_Tbl()
					_agda := wml.NewEG_BlockLevelElts()
					_fbdg := wml.NewEG_ContentBlockContent()
					_agda.EG_ContentBlockContent = append(_agda.EG_ContentBlockContent, _fbdg)
					_fbdg.Tbl = append(_fbdg.Tbl, _ada)
					_efe.EG_BlockLevelElts = append(_efe.EG_BlockLevelElts, nil)
					if _gcc {
						copy(_efe.EG_BlockLevelElts[_cfa+1:], _efe.EG_BlockLevelElts[_cfa:])
						_efe.EG_BlockLevelElts[_cfa] = _agda
						if _ege != 0 {
							_afg := wml.NewEG_BlockLevelElts()
							_bgg := wml.NewEG_ContentBlockContent()
							_afg.EG_ContentBlockContent = append(_afg.EG_ContentBlockContent, _bgg)
							_bgg.P = _gdc.P[:_ege]
							_efe.EG_BlockLevelElts = append(_efe.EG_BlockLevelElts, nil)
							copy(_efe.EG_BlockLevelElts[_cfa+1:], _efe.EG_BlockLevelElts[_cfa:])
							_efe.EG_BlockLevelElts[_cfa] = _afg
						}
						_gdc.P = _gdc.P[_ege:]
					} else {
						copy(_efe.EG_BlockLevelElts[_cfa+2:], _efe.EG_BlockLevelElts[_cfa+1:])
						_efe.EG_BlockLevelElts[_cfa+1] = _agda
						if _ege != len(_gdc.P)-1 {
							_dfa := wml.NewEG_BlockLevelElts()
							_bde := wml.NewEG_ContentBlockContent()
							_dfa.EG_ContentBlockContent = append(_dfa.EG_ContentBlockContent, _bde)
							_bde.P = _gdc.P[_ege+1:]
							_efe.EG_BlockLevelElts = append(_efe.EG_BlockLevelElts, nil)
							copy(_efe.EG_BlockLevelElts[_cfa+3:], _efe.EG_BlockLevelElts[_cfa+2:])
							_efe.EG_BlockLevelElts[_cfa+2] = _dfa
						}
						_gdc.P = _gdc.P[:_ege+1]
					}
					return Table{_bba, _ada}
				}
			}
			for _, _add := range _gdc.Tbl {
				_fcb := _adaa(_add, _cbd, _gcc)
				if _fcb != nil {
					return Table{_bba, _fcb}
				}
			}
		}
	}
	return _bba.AddTable()
}

// Clear clears the styes.
func (_eccb Styles) Clear() {
	_eccb._abca.DocDefaults = nil
	_eccb._abca.LatentStyles = nil
	_eccb._abca.Style = nil
}

// SetContextualSpacing controls whether to Ignore Spacing Above and Below When
// Using Identical Styles
func (_cfdgf ParagraphStyleProperties) SetContextualSpacing(b bool) {
	if !b {
		_cfdgf._gfee.ContextualSpacing = nil
	} else {
		_cfdgf._gfee.ContextualSpacing = wml.NewCT_OnOff()
	}
}

// Underline returns the type of run underline.
func (_fdde RunProperties) Underline() wml.ST_Underline {
	if _febb := _fdde._gbdb.U; _febb != nil {
		return _febb.ValAttr
	}
	return 0
}

// SetAfterLineSpacing sets spacing below paragraph in line units.
func (_dfac Paragraph) SetAfterLineSpacing(d measurement.Distance) {
	_dfac.ensurePPr()
	if _dfac._eagd.PPr.Spacing == nil {
		_dfac._eagd.PPr.Spacing = wml.NewCT_Spacing()
	}
	_bcgf := _dfac._eagd.PPr.Spacing
	_bcgf.AfterLinesAttr = unioffice.Int64(int64(d / measurement.Twips))
}

// AddParagraph adds a paragraph to the endnote.
func (_bdeg Endnote) AddParagraph() Paragraph {
	_gbbc := wml.NewEG_ContentBlockContent()
	_bcgag := len(_bdeg._fagg.EG_BlockLevelElts[0].EG_ContentBlockContent)
	_bdeg._fagg.EG_BlockLevelElts[0].EG_ContentBlockContent = append(_bdeg._fagg.EG_BlockLevelElts[0].EG_ContentBlockContent, _gbbc)
	_aea := wml.NewCT_P()
	var _dcgbb *wml.CT_String
	if _bcgag != 0 {
		_bffga := len(_bdeg._fagg.EG_BlockLevelElts[0].EG_ContentBlockContent[_bcgag-1].P)
		_dcgbb = _bdeg._fagg.EG_BlockLevelElts[0].EG_ContentBlockContent[_bcgag-1].P[_bffga-1].PPr.PStyle
	} else {
		_dcgbb = wml.NewCT_String()
		_dcgbb.ValAttr = "\u0045n\u0064\u006e\u006f\u0074\u0065"
	}
	_gbbc.P = append(_gbbc.P, _aea)
	_agfgb := Paragraph{_bdeg._cceg, _aea}
	_agfgb._eagd.PPr = wml.NewCT_PPr()
	_agfgb._eagd.PPr.PStyle = _dcgbb
	_agfgb._eagd.PPr.RPr = wml.NewCT_ParaRPr()
	return _agfgb
}

// SetAfterSpacing sets spacing below paragraph.
func (_cfdf Paragraph) SetAfterSpacing(d measurement.Distance) {
	_cfdf.ensurePPr()
	if _cfdf._eagd.PPr.Spacing == nil {
		_cfdf._eagd.PPr.Spacing = wml.NewCT_Spacing()
	}
	_bfbec := _cfdf._eagd.PPr.Spacing
	_bfbec.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
	_bfbec.AfterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(d / measurement.Twips))
}

// SetName sets the name of the style.
func (_cgef Style) SetName(name string) {
	_cgef._gaege.Name = wml.NewCT_String()
	_cgef._gaege.Name.ValAttr = name
}

// Font returns the name of run font family.
func (_gaaee RunProperties) Font() string {
	if _fccc := _gaaee._gbdb.RFonts; _fccc != nil {
		if _fccc.AsciiAttr != nil {
			return *_fccc.AsciiAttr
		} else if _fccc.HAnsiAttr != nil {
			return *_fccc.HAnsiAttr
		} else if _fccc.CsAttr != nil {
			return *_fccc.CsAttr
		}
	}
	return ""
}

// SetColor sets a specific color or auto.
func (_bag Color) SetColor(v color.Color) {
	if v.IsAuto() {
		_bag._ec.ValAttr.ST_HexColorAuto = wml.ST_HexColorAutoAuto
		_bag._ec.ValAttr.ST_HexColorRGB = nil
	} else {
		_bag._ec.ValAttr.ST_HexColorAuto = wml.ST_HexColorAutoUnset
		_bag._ec.ValAttr.ST_HexColorRGB = v.AsRGBString()
	}
}

// ParagraphProperties are the properties for a paragraph.
type ParagraphProperties struct {
	_aage *Document
	_dfaf *wml.CT_PPr
}

// DocText is an array of extracted text items which has some methods for representing extracted text.
type DocText struct {
	Items []TextItem
	_aefd []listItemInfo
	_fddc map[int64]map[int64]int64
}

// Settings controls the document settings.
type Settings struct{ _cdbbf *wml.Settings }

// GetColor returns the color.Color object representing the run color.
func (_abggf ParagraphProperties) GetColor() color.Color {
	if _ddgg := _abggf._dfaf.RPr.Color; _ddgg != nil {
		_gafbe := _ddgg.ValAttr
		if _gafbe.ST_HexColorRGB != nil {
			return color.FromHex(*_gafbe.ST_HexColorRGB)
		}
	}
	return color.Color{}
}

// read reads a document from an io.Reader.
func Read(r io.ReaderAt, size int64) (*Document, error) { return _fbee(r, size, "") }

// SetName sets the name of the bookmark. This is the name that is used to
// reference the bookmark from hyperlinks.
func (_eecf Bookmark) SetName(name string) { _eecf._gc.NameAttr = name }

// Runs returns all of the runs in a paragraph.
func (_bgbac Paragraph) Runs() []Run {
	_fgge := []Run{}
	for _, _ebdge := range _bgbac._eagd.EG_PContent {
		if _ebdge.Hyperlink != nil && _ebdge.Hyperlink.EG_ContentRunContent != nil {
			for _, _dcggc := range _ebdge.Hyperlink.EG_ContentRunContent {
				if _dcggc.R != nil {
					_fgge = append(_fgge, Run{_bgbac._fagf, _dcggc.R})
				}
			}
		}
		for _, _abcf := range _ebdge.EG_ContentRunContent {
			if _abcf.R != nil {
				_fgge = append(_fgge, Run{_bgbac._fagf, _abcf.R})
			}
			if _abcf.Sdt != nil && _abcf.Sdt.SdtContent != nil {
				for _, _bcec := range _abcf.Sdt.SdtContent.EG_ContentRunContent {
					if _bcec.R != nil {
						_fgge = append(_fgge, Run{_bgbac._fagf, _bcec.R})
					}
				}
			}
		}
	}
	return _fgge
}

// X returns the inner wrapped XML type.
func (_dadad TableStyleProperties) X() *wml.CT_TblPrBase { return _dadad._degc }

// Bold returns true if paragraph font is bold.
func (_dgcbb ParagraphProperties) Bold() bool {
	_edfdb := _dgcbb._dfaf.RPr
	return _cadf(_edfdb.B) || _cadf(_edfdb.BCs)
}
func (_eaag Paragraph) addEndFldChar() *wml.CT_FldChar {
	_fega := _eaag.addFldChar()
	_fega.FldCharTypeAttr = wml.ST_FldCharTypeEnd
	return _fega
}
func (_bgf *Document) save(_bad io.Writer, _dcc string) error {
	const _baa = "\u0064o\u0063u\u006d\u0065\u006e\u0074\u003a\u0064\u002e\u0053\u0061\u0076\u0065"
	if _gad := _bgf.doc.Validate(); _gad != nil {
		logger.Log.Warning("\u0076\u0061\u006c\u0069\u0064\u0061\u0074\u0069\u006f\u006e\u0020\u0065\u0072\u0072\u006fr\u0020i\u006e\u0020\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u003a\u0020\u0025\u0073", _gad)
	}
	_dbg := unioffice.DocTypeDocument
	if !license.GetLicenseKey().IsLicensed() && !_eece {
		fmt.Println("\u0055\u006e\u006ci\u0063\u0065\u006e\u0073e\u0064\u0020\u0076\u0065\u0072\u0073\u0069o\u006e\u0020\u006f\u0066\u0020\u0055\u006e\u0069\u004f\u0066\u0066\u0069\u0063\u0065")
		fmt.Println("\u002d\u0020\u0047e\u0074\u0020\u0061\u0020\u0074\u0072\u0069\u0061\u006c\u0020\u006c\u0069\u0063\u0065\u006e\u0073\u0065\u0020\u006f\u006e\u0020\u0068\u0074\u0074\u0070\u0073\u003a\u002f\u002fu\u006e\u0069\u0064\u006f\u0063\u002e\u0069\u006f")
		return errors.New("\u0075\u006e\u0069\u006f\u0066\u0066\u0069\u0063\u0065\u0020\u006ci\u0063\u0065\u006e\u0073\u0065\u0020\u0072\u0065\u0071\u0075i\u0072\u0065\u0064")
	}
	if len(_bgf._feg) == 0 {
		if len(_dcc) > 0 {
			_bgf._feg = _dcc
		} else {
			_cca, _beaf := license.GenRefId("\u0064\u0077")
			if _beaf != nil {
				logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _beaf)
				return _beaf
			}
			_bgf._feg = _cca
		}
	}
	if _agcd := license.Track(_bgf._feg, _baa); _agcd != nil {
		logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _agcd)
		return _agcd
	}
	_dccg := zip.NewWriter(_bad)
	defer _dccg.Close()
	if _cfe := zippkg.MarshalXML(_dccg, unioffice.BaseRelsFilename, _bgf.Rels.X()); _cfe != nil {
		return _cfe
	}
	if _bbdg := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.ExtendedPropertiesType, _bgf.AppProperties.X()); _bbdg != nil {
		return _bbdg
	}
	if _caec := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.CorePropertiesType, _bgf.CoreProperties.X()); _caec != nil {
		return _caec
	}
	if _bgf.CustomProperties.X() != nil {
		if _fbdb := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.CustomPropertiesType, _bgf.CustomProperties.X()); _fbdb != nil {
			return _fbdb
		}
	}
	if _bgf.Thumbnail != nil {
		_fce, _fea := _dccg.Create("\u0064\u006f\u0063Pr\u006f\u0070\u0073\u002f\u0074\u0068\u0075\u006d\u0062\u006e\u0061\u0069\u006c\u002e\u006a\u0070\u0065\u0067")
		if _fea != nil {
			return _fea
		}
		if _fdc := jpeg.Encode(_fce, _bgf.Thumbnail, nil); _fdc != nil {
			return _fdc
		}
	}
	if _bgd := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.SettingsType, _bgf.Settings.X()); _bgd != nil {
		return _bgd
	}
	_acde := unioffice.AbsoluteFilename(_dbg, unioffice.OfficeDocumentType, 0)
	if _ecd := zippkg.MarshalXML(_dccg, _acde, _bgf.doc); _ecd != nil {
		return _ecd
	}
	if _dada := zippkg.MarshalXML(_dccg, zippkg.RelationsPathFor(_acde), _bgf._dab.X()); _dada != nil {
		return _dada
	}
	if _bgf.Numbering.X() != nil {
		if _aeda := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.NumberingType, _bgf.Numbering.X()); _aeda != nil {
			return _aeda
		}
	}
	if _gcf := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.StylesType, _bgf.Styles.X()); _gcf != nil {
		return _gcf
	}
	if _bgf._gbe != nil {
		if _aeg := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.WebSettingsType, _bgf._gbe); _aeg != nil {
			return _aeg
		}
	}
	if _bgf._eaa != nil {
		if _fege := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.FontTableType, _bgf._eaa); _fege != nil {
			return _fege
		}
	}
	if _bgf._ccb != nil {
		if _cdag := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.EndNotesType, _bgf._ccb); _cdag != nil {
			return _cdag
		}
	}
	if _bgf._beg != nil {
		if _cegf := zippkg.MarshalXMLByType(_dccg, _dbg, unioffice.FootNotesType, _bgf._beg); _cegf != nil {
			return _cegf
		}
	}
	for _gbbe, _efg := range _bgf._ffbc {
		if _ffbb := zippkg.MarshalXMLByTypeIndex(_dccg, _dbg, unioffice.ThemeType, _gbbe+1, _efg); _ffbb != nil {
			return _ffbb
		}
	}
	for _abad, _gcfb := range _bgf._gga {
		_fead, _fee := _gcfb.ExportToByteArray()
		if _fee != nil {
			return _fee
		}
		_fgf := "\u0077\u006f\u0072d\u002f" + _gcfb.TargetAttr[:len(_gcfb.TargetAttr)-4] + "\u002e\u0062\u0069\u006e"
		if _cdf := zippkg.AddFileFromBytes(_dccg, _fgf, _fead); _cdf != nil {
			return _cdf
		}
		if _ebf := zippkg.MarshalXMLByTypeIndex(_dccg, _dbg, unioffice.ControlType, _abad+1, _gcfb.Ocx); _ebf != nil {
			return _ebf
		}
	}
	for _fdb, _aec := range _bgf._geb {
		_dgcg := unioffice.AbsoluteFilename(_dbg, unioffice.HeaderType, _fdb+1)
		if _bga := zippkg.MarshalXML(_dccg, _dgcg, _aec); _bga != nil {
			return _bga
		}
		if !_bgf._cbfd[_fdb].IsEmpty() {
			zippkg.MarshalXML(_dccg, zippkg.RelationsPathFor(_dgcg), _bgf._cbfd[_fdb].X())
		}
	}
	for _ade, _bff := range _bgf._aba {
		_dgf := unioffice.AbsoluteFilename(_dbg, unioffice.FooterType, _ade+1)
		if _fag := zippkg.MarshalXMLByTypeIndex(_dccg, _dbg, unioffice.FooterType, _ade+1, _bff); _fag != nil {
			return _fag
		}
		if !_bgf._fdf[_ade].IsEmpty() {
			zippkg.MarshalXML(_dccg, zippkg.RelationsPathFor(_dgf), _bgf._fdf[_ade].X())
		}
	}
	for _dda, _dbc := range _bgf.Images {
		if _bca := common.AddImageToZip(_dccg, _dbc, _dda+1, unioffice.DocTypeDocument); _bca != nil {
			return _bca
		}
	}
	for _fdg, _cac := range _bgf._caf {
		_ggc := unioffice.AbsoluteFilename(_dbg, unioffice.ChartType, _fdg+1)
		zippkg.MarshalXML(_dccg, _ggc, _cac._ffb)
	}
	if _bcc := zippkg.MarshalXML(_dccg, unioffice.ContentTypesFilename, _bgf.ContentTypes.X()); _bcc != nil {
		return _bcc
	}
	if _bbce := _bgf.WriteExtraFiles(_dccg); _bbce != nil {
		return _bbce
	}
	return _dccg.Close()
}

// Document is a text document that can be written out in the OOXML .docx
// format. It can be opened from a file on disk and modified, or created from
// scratch.
type Document struct {
	common.DocBase
	doc      *wml.Document
	Settings  Settings
	Numbering Numbering
	Styles    Styles
	_geb      []*wml.Hdr
	_cbfd     []common.Relationships
	_aba      []*wml.Ftr
	_fdf      []common.Relationships
	_dab      common.Relationships
	_ffbc     []*dml.Theme
	_gbe      *wml.WebSettings
	_eaa      *wml.Fonts
	_ccb      *wml.Endnotes
	_beg      *wml.Footnotes
	_gga      []*axcontrol.Control
	_caf      []*chart
	_feg      string
}

// NewAnchorDrawWrapOptions return anchor drawing options property.
func NewAnchorDrawWrapOptions() *AnchorDrawWrapOptions {
	_ba := &AnchorDrawWrapOptions{}
	if !_ba._cef {
		_gd, _bd := _afa()
		_ba._dd = _gd
		_ba._cbf = _bd
	}
	return _ba
}

// TextWithOptions extract text with options.
func (_bgcb *DocText) TextWithOptions(options ExtractTextOptions) string {
	_affc := make(map[int64]map[int64]int64, 0)
	_fba := bytes.NewBuffer([]byte{})
	_bcaff := int64(0)
	_egcg := int64(0)
	_gacbf := int64(0)
	for _bcdee, _feaa := range _bgcb.Items {
		_eae := false
		if _feaa.Text != "" {
			if options.WithNumbering {
				if _bcdee > 0 {
					if _feaa.Paragraph != _bgcb.Items[_bcdee-1].Paragraph {
						_eae = true
					}
				} else {
					_eae = true
				}
				if _eae {
					for _, _dedb := range _bgcb._aefd {
						if _dedb.FromParagraph == nil {
							continue
						}
						if _dedb.FromParagraph.X() == _feaa.Paragraph {
							if _bdegf := _dedb.NumberingLevel.X(); _bdegf != nil {
								if _dedb.AbstractNumId != nil && _bgcb._fddc[*_dedb.AbstractNumId][_bdegf.IlvlAttr] > 0 {
									if _, _aaegf := _affc[*_dedb.AbstractNumId]; _aaegf {
										if _, _dfagd := _affc[*_dedb.AbstractNumId][_bdegf.IlvlAttr]; _dfagd {
											_affc[*_dedb.AbstractNumId][_bdegf.IlvlAttr]++
										} else {
											_affc[*_dedb.AbstractNumId][_bdegf.IlvlAttr] = 1
										}
									} else {
										_affc[*_dedb.AbstractNumId] = map[int64]int64{_bdegf.IlvlAttr: 1}
									}
									if _bcaff == _dedb.NumberingLevel.X().IlvlAttr && _bdegf.IlvlAttr > 0 {
										_egcg++
									} else {
										_egcg = _affc[*_dedb.AbstractNumId][_bdegf.IlvlAttr]
										if _bdegf.IlvlAttr > _bcaff && _gacbf == *_dedb.AbstractNumId {
											_egcg = 1
										}
									}
									_afbgb := ""
									if _bdegf.LvlText.ValAttr != nil {
										_afbgb = *_bdegf.LvlText.ValAttr
									}
									_face := formatutils.FormatNumberingText(_egcg, _bdegf.IlvlAttr, _afbgb, _bdegf.NumFmt, _affc[*_dedb.AbstractNumId])
									_fba.WriteString(_face)
									_bgcb._fddc[*_dedb.AbstractNumId][_bdegf.IlvlAttr]--
									_bcaff = _bdegf.IlvlAttr
									_gacbf = *_dedb.AbstractNumId
									if options.NumberingIndent != "" {
										_fba.WriteString(options.NumberingIndent)
									}
								}
							}
							break
						}
					}
				}
			}
			_fba.WriteString(_feaa.Text)
			_fba.WriteString("\u000a")
		}
	}
	return _fba.String()
}

// SetOrigin sets the origin of the image.  It defaults to ST_RelFromHPage and
// ST_RelFromVPage
func (_bg AnchoredDrawing) SetOrigin(h wml.WdST_RelFromH, v wml.WdST_RelFromV) {
	_bg._dgc.PositionH.RelativeFromAttr = h
	_bg._dgc.PositionV.RelativeFromAttr = v
}

// X return slice of node.
func (_dfed *Nodes) X() []Node { return _dfed._gabfc }

// NewTableWidth returns a newly intialized TableWidth
func NewTableWidth() TableWidth { return TableWidth{wml.NewCT_TblWidth()} }
func (_efdg Paragraph) insertRun(_defd Run, _dbcb bool) Run {
	for _, _fgggb := range _efdg._eagd.EG_PContent {
		for _bbeb, _fgba := range _fgggb.EG_ContentRunContent {
			if _fgba.R == _defd.X() {
				_eafe := wml.NewCT_R()
				_fgggb.EG_ContentRunContent = append(_fgggb.EG_ContentRunContent, nil)
				if _dbcb {
					copy(_fgggb.EG_ContentRunContent[_bbeb+1:], _fgggb.EG_ContentRunContent[_bbeb:])
					_fgggb.EG_ContentRunContent[_bbeb] = wml.NewEG_ContentRunContent()
					_fgggb.EG_ContentRunContent[_bbeb].R = _eafe
				} else {
					copy(_fgggb.EG_ContentRunContent[_bbeb+2:], _fgggb.EG_ContentRunContent[_bbeb+1:])
					_fgggb.EG_ContentRunContent[_bbeb+1] = wml.NewEG_ContentRunContent()
					_fgggb.EG_ContentRunContent[_bbeb+1].R = _eafe
				}
				return Run{_efdg._fagf, _eafe}
			}
			if _fgba.Sdt != nil && _fgba.Sdt.SdtContent != nil {
				for _, _eddb := range _fgba.Sdt.SdtContent.EG_ContentRunContent {
					if _eddb.R == _defd.X() {
						_bgfc := wml.NewCT_R()
						_fgba.Sdt.SdtContent.EG_ContentRunContent = append(_fgba.Sdt.SdtContent.EG_ContentRunContent, nil)
						if _dbcb {
							copy(_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb+1:], _fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb:])
							_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb] = wml.NewEG_ContentRunContent()
							_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb].R = _bgfc
						} else {
							copy(_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb+2:], _fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb+1:])
							_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb+1] = wml.NewEG_ContentRunContent()
							_fgba.Sdt.SdtContent.EG_ContentRunContent[_bbeb+1].R = _bgfc
						}
						return Run{_efdg._fagf, _bgfc}
					}
				}
			}
		}
	}
	return _efdg.AddRun()
}

// AddEndnote will create a new endnote and attach it to the Paragraph in the
// location at the end of the previous run (endnotes create their own run within
// the paragraph. The text given to the function is simply a convenience helper,
// paragraphs and runs can always be added to the text of the endnote later.
func (_aeff Paragraph) AddEndnote(text string) Endnote {
	var _cegac int64
	if _aeff._fagf.HasEndnotes() {
		for _, _dbdf := range _aeff._fagf.Endnotes() {
			if _dbdf.id() > _cegac {
				_cegac = _dbdf.id()
			}
		}
		_cegac++
	} else {
		_cegac = 0
		_aeff._fagf._ccb = &wml.Endnotes{}
	}
	_deca := wml.NewCT_FtnEdn()
	_bfge := wml.NewCT_FtnEdnRef()
	_bfge.IdAttr = _cegac
	_aeff._fagf._ccb.CT_Endnotes.Endnote = append(_aeff._fagf._ccb.CT_Endnotes.Endnote, _deca)
	_fecgec := _aeff.AddRun()
	_eefac := _fecgec.Properties()
	_eefac.SetStyle("\u0045\u006e\u0064\u006e\u006f\u0074\u0065\u0041\u006e\u0063\u0068\u006f\u0072")
	_fecgec._adaad.EG_RunInnerContent = []*wml.EG_RunInnerContent{wml.NewEG_RunInnerContent()}
	_fecgec._adaad.EG_RunInnerContent[0].EndnoteReference = _bfge
	_abbcgb := Endnote{_aeff._fagf, _deca}
	_abbcgb._fagg.IdAttr = _cegac
	_abbcgb._fagg.EG_BlockLevelElts = []*wml.EG_BlockLevelElts{wml.NewEG_BlockLevelElts()}
	_dcfa := _abbcgb.AddParagraph()
	_dcfa.Properties().SetStyle("\u0045n\u0064\u006e\u006f\u0074\u0065")
	_dcfa._eagd.PPr.RPr = wml.NewCT_ParaRPr()
	_egfd := _dcfa.AddRun()
	_egfd.AddTab()
	_egfd.AddText(text)
	return _abbcgb
}

// Rows returns the rows defined in the table.
func (_dabcb Table) Rows() []Row {
	_ebeac := []Row{}
	for _, _abgcf := range _dabcb.ctTbl.EG_ContentRowContent {
		for _, _baege := range _abgcf.Tr {
			_ebeac = append(_ebeac, Row{_dabcb.doc, _baege})
		}
		if _abgcf.Sdt != nil && _abgcf.Sdt.SdtContent != nil {
			for _, _dfgb := range _abgcf.Sdt.SdtContent.Tr {
				_ebeac = append(_ebeac, Row{_dabcb.doc, _dfgb})
			}
		}
	}
	return _ebeac
}

// Paragraph is a paragraph within a document.
type Paragraph struct {
	_fagf *Document
	_eagd *wml.CT_P
}

// SetStartIndent controls the start indent of the paragraph.
func (_afgaf ParagraphStyleProperties) SetStartIndent(m measurement.Distance) {
	if _afgaf._gfee.Ind == nil {
		_afgaf._gfee.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_afgaf._gfee.Ind.StartAttr = nil
	} else {
		_afgaf._gfee.Ind.StartAttr = &wml.ST_SignedTwipsMeasure{}
		_afgaf._gfee.Ind.StartAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// X return element of Node as interface, can be either *Paragraph, *Table and Run.
func (_dcea *Node) X() interface{} { return _dcea._ggda }

// Borders allows manipulation of the table borders.
func (_cbgag TableStyleProperties) Borders() TableBorders {
	if _cbgag._degc.TblBorders == nil {
		_cbgag._degc.TblBorders = wml.NewCT_TblBorders()
	}
	return TableBorders{_cbgag._degc.TblBorders}
}

// ReplaceTextByRegexp replace text inside node using regexp.
func (_dbfe *Nodes) ReplaceTextByRegexp(expr *regexp.Regexp, newText string) {
	for _, _gaba := range _dbfe._gabfc {
		_gaba.ReplaceTextByRegexp(expr, newText)
	}
}
func _eeebf(_fcgg *wml.CT_P, _dggd *wml.CT_Hyperlink, _bdgf *TableInfo, _gfbc *DrawingInfo, _dcca []*wml.EG_PContent) []TextItem {
	if len(_dcca) == 0 {
		return []TextItem{TextItem{Text: "", DrawingInfo: _gfbc, Paragraph: _fcgg, Hyperlink: _dggd, Run: nil, TableInfo: _bdgf}}
	}
	_fdbb := []TextItem{}
	for _, _fbfb := range _dcca {
		for _, _abcb := range _fbfb.FldSimple {
			if _abcb != nil {
				_fdbb = append(_fdbb, _eeebf(_fcgg, _dggd, _bdgf, _gfbc, _abcb.EG_PContent)...)
			}
		}
		if _fffc := _fbfb.Hyperlink; _fffc != nil {
			_fdbb = append(_fdbb, _bfafd(_fcgg, _fffc, _bdgf, _gfbc, _fffc.EG_ContentRunContent)...)
		}
		_fdbb = append(_fdbb, _bfafd(_fcgg, nil, _bdgf, _gfbc, _fbfb.EG_ContentRunContent)...)
	}
	return _fdbb
}

// AddField adds a field (automatically computed text) to the document.
func (_accbc Run) AddField(code string) { _accbc.AddFieldWithFormatting(code, "", true) }

// Append appends a document d0 to a document d1. All settings, headers and footers remain the same as in the document d0 if they exist there, otherwise they are taken from the d1.
func (_beag *Document) Append(d1orig *Document) error {
	_dcbgf, _bgcd := d1orig.Copy()
	if _bgcd != nil {
		return _bgcd
	}
	_beag.DocBase = _beag.DocBase.Append(_dcbgf.DocBase)
	if _dcbgf.doc.ConformanceAttr != sharedTypes.ST_ConformanceClassStrict {
		_beag.doc.ConformanceAttr = _dcbgf.doc.ConformanceAttr
	}
	_abfc := _beag._dab.X().Relationship
	_cee := _dcbgf._dab.X().Relationship
	_geed := _dcbgf.doc.Body
	_degb := map[string]string{}
	_facf := map[int64]int64{}
	_abdac := map[int64]int64{}
	for _, _accd := range _cee {
		_dagf := true
		_daae := _accd.IdAttr
		_bega := _accd.TargetAttr
		_bgda := _accd.TypeAttr
		_accf := _bgda == unioffice.ImageType
		_bfdc := _bgda == unioffice.HyperLinkType
		var _ddace string
		for _, _daee := range _abfc {
			if _daee.TypeAttr == _bgda && _daee.TargetAttr == _bega {
				_dagf = false
				_ddace = _daee.IdAttr
				break
			}
		}
		if _accf {
			_dbfa := "\u0077\u006f\u0072d\u002f" + _bega
			for _, _bafcc := range _dcbgf.DocBase.Images {
				if _bafcc.Target() == _dbfa {
					_abga, _cbbf := common.ImageFromStorage(_bafcc.Path())
					if _cbbf != nil {
						return _cbbf
					}
					_gccc, _cbbf := _beag.AddImage(_abga)
					if _cbbf != nil {
						return _cbbf
					}
					_ddace = _gccc.RelID()
					break
				}
			}
		} else if _dagf {
			if _bfdc {
				_adabb := _beag._dab.AddHyperlink(_bega)
				_ddace = common.Relationship(_adabb).ID()
			} else {
				_dfbf := _beag._dab.AddRelationship(_bega, _bgda)
				_ddace = _dfbf.X().IdAttr
			}
		}
		if _daae != _ddace {
			_degb[_daae] = _ddace
		}
	}
	if _geed.SectPr != nil {
		for _, _adba := range _geed.SectPr.EG_HdrFtrReferences {
			if _adba.HeaderReference != nil {
				if _fcae, _bggf := _degb[_adba.HeaderReference.IdAttr]; _bggf {
					_adba.HeaderReference.IdAttr = _fcae
					_beag._cbfd = append(_beag._cbfd, common.NewRelationships())
				}
			} else if _adba.FooterReference != nil {
				if _dffd, _ebc := _degb[_adba.FooterReference.IdAttr]; _ebc {
					_adba.FooterReference.IdAttr = _dffd
					_beag._fdf = append(_beag._fdf, common.NewRelationships())
				}
			}
		}
	}
	_dfe, _efdfb := _beag._ccb, _dcbgf._ccb
	if _dfe != nil {
		if _efdfb != nil {
			if _dfe.Endnote != nil {
				if _efdfb.Endnote != nil {
					_bdfbc := int64(len(_dfe.Endnote) + 1)
					for _, _abbb := range _efdfb.Endnote {
						_abbcf := _abbb.IdAttr
						if _abbcf > 0 {
							_abbb.IdAttr = _bdfbc
							_dfe.Endnote = append(_dfe.Endnote, _abbb)
							_abdac[_abbcf] = _bdfbc
							_bdfbc++
						}
					}
				}
			} else {
				_dfe.Endnote = _efdfb.Endnote
			}
		}
	} else if _efdfb != nil {
		_dfe = _efdfb
	}
	_beag._ccb = _dfe
	_abgg, _acdd := _beag._beg, _dcbgf._beg
	if _abgg != nil {
		if _acdd != nil {
			if _abgg.Footnote != nil {
				if _acdd.Footnote != nil {
					_ebfg := int64(len(_abgg.Footnote) + 1)
					for _, _gbada := range _acdd.Footnote {
						_aeea := _gbada.IdAttr
						if _aeea > 0 {
							_gbada.IdAttr = _ebfg
							_abgg.Footnote = append(_abgg.Footnote, _gbada)
							_facf[_aeea] = _ebfg
							_ebfg++
						}
					}
				}
			} else {
				_abgg.Footnote = _acdd.Footnote
			}
		}
	} else if _acdd != nil {
		_abgg = _acdd
	}
	_beag._beg = _abgg
	for _, _badd := range _geed.EG_BlockLevelElts {
		for _, _eeec := range _badd.EG_ContentBlockContent {
			for _, _bfgf := range _eeec.P {
				_bfgge(_bfgf, _degb)
				_cbdfg(_bfgf, _degb)
				_bfgff(_bfgf, _facf, _abdac)
			}
			for _, _cbdg := range _eeec.Tbl {
				_aefc(_cbdg, _degb)
				_gcdg(_cbdg, _degb)
				_eebg(_cbdg, _facf, _abdac)
			}
		}
	}
	_beag.doc.Body.EG_BlockLevelElts = append(_beag.doc.Body.EG_BlockLevelElts, _dcbgf.doc.Body.EG_BlockLevelElts...)
	if _beag.doc.Body.SectPr == nil {
		_beag.doc.Body.SectPr = _dcbgf.doc.Body.SectPr
	} else {
		var _afge, _bfbd bool
		for _, _dgdg := range _beag.doc.Body.SectPr.EG_HdrFtrReferences {
			if _dgdg.HeaderReference != nil {
				_afge = true
			} else if _dgdg.FooterReference != nil {
				_bfbd = true
			}
		}
		if !_afge {
			for _, _fdd := range _dcbgf.doc.Body.SectPr.EG_HdrFtrReferences {
				if _fdd.HeaderReference != nil {
					_beag.doc.Body.SectPr.EG_HdrFtrReferences = append(_beag.doc.Body.SectPr.EG_HdrFtrReferences, _fdd)
					break
				}
			}
		}
		if !_bfbd {
			for _, _dbaef := range _dcbgf.doc.Body.SectPr.EG_HdrFtrReferences {
				if _dbaef.FooterReference != nil {
					_beag.doc.Body.SectPr.EG_HdrFtrReferences = append(_beag.doc.Body.SectPr.EG_HdrFtrReferences, _dbaef)
					break
				}
			}
		}
		if _beag.doc.Body.SectPr.Cols == nil && _dcbgf.doc.Body.SectPr.Cols != nil {
			_beag.doc.Body.SectPr.Cols = _dcbgf.doc.Body.SectPr.Cols
		}
	}
	_ceca := _beag.Numbering._cbag
	_babd := _dcbgf.Numbering._cbag
	if _ceca != nil {
		if _babd != nil {
			_ceca.NumPicBullet = append(_ceca.NumPicBullet, _babd.NumPicBullet...)
			_ceca.AbstractNum = append(_ceca.AbstractNum, _babd.AbstractNum...)
			_ceca.Num = append(_ceca.Num, _babd.Num...)
		}
	} else if _babd != nil {
		_ceca = _babd
	}
	_beag.Numbering._cbag = _ceca
	if _beag.Styles._abca == nil && _dcbgf.Styles._abca != nil {
		_beag.Styles._abca = _dcbgf.Styles._abca
	}
	_beag._ffbc = append(_beag._ffbc, _dcbgf._ffbc...)
	_beag._gga = append(_beag._gga, _dcbgf._gga...)
	if len(_beag._geb) == 0 {
		_beag._geb = _dcbgf._geb
	}
	if len(_beag._aba) == 0 {
		_beag._aba = _dcbgf._aba
	}
	_efbd := _beag._gbe
	_fded := _dcbgf._gbe
	if _efbd != nil {
		if _fded != nil {
			if _efbd.Divs != nil {
				if _fded.Divs != nil {
					_efbd.Divs.Div = append(_efbd.Divs.Div, _fded.Divs.Div...)
				}
			} else {
				_efbd.Divs = _fded.Divs
			}
		}
		_efbd.Frameset = nil
	} else if _fded != nil {
		_efbd = _fded
		_efbd.Frameset = nil
	}
	_beag._gbe = _efbd
	_dacbf := _beag._eaa
	_acbf := _dcbgf._eaa
	if _dacbf != nil {
		if _acbf != nil {
			if _dacbf.Font != nil {
				if _acbf.Font != nil {
					for _, _gaed := range _acbf.Font {
						_fgfg := true
						for _, _caccf := range _dacbf.Font {
							if _caccf.NameAttr == _gaed.NameAttr {
								_fgfg = false
								break
							}
						}
						if _fgfg {
							_dacbf.Font = append(_dacbf.Font, _gaed)
						}
					}
				}
			} else {
				_dacbf.Font = _acbf.Font
			}
		}
	} else if _acbf != nil {
		_dacbf = _acbf
	}
	_beag._eaa = _dacbf
	return nil
}

// SetCharacterSpacing sets the run's Character Spacing Adjustment.
func (_egbf RunProperties) SetCharacterSpacing(size measurement.Distance) {
	_egbf._gbdb.Spacing = wml.NewCT_SignedTwipsMeasure()
	_egbf._gbdb.Spacing.ValAttr.Int64 = unioffice.Int64(int64(size / measurement.Twips))
}

// AddImageRef add ImageRef to header as relationship, returning ImageRef
// that can be used to be placed as header content.
func (_gcab Header) AddImageRef(r common.ImageRef) (common.ImageRef, error) {
	var _fcfg common.Relationships
	for _daef, _gbbf := range _gcab._dbagd._geb {
		if _gbbf == _gcab._deae {
			_fcfg = _gcab._dbagd._cbfd[_daef]
		}
	}
	_gcbb := _fcfg.AddRelationship(r.Target(), unioffice.ImageType)
	r.SetRelID(_gcbb.X().IdAttr)
	return r, nil
}
func (_eeb *Document) appendParagraph(_afdfa *Paragraph, _ecgb Paragraph, _dcbg bool) Paragraph {
	_cdga := wml.NewEG_BlockLevelElts()
	_eeb.doc.Body.EG_BlockLevelElts = append(_eeb.doc.Body.EG_BlockLevelElts, _cdga)
	_adfe := wml.NewEG_ContentBlockContent()
	_cdga.EG_ContentBlockContent = append(_cdga.EG_ContentBlockContent, _adfe)
	if _afdfa != nil {
		_bbec := _afdfa.X()
		for _, _bafa := range _eeb.doc.Body.EG_BlockLevelElts {
			for _, _egda := range _bafa.EG_ContentBlockContent {
				for _ede, _gfef := range _egda.P {
					if _gfef == _bbec {
						_fdcg := _ecgb.X()
						_egda.P = append(_egda.P, nil)
						if _dcbg {
							copy(_egda.P[_ede+1:], _egda.P[_ede:])
							_egda.P[_ede] = _fdcg
						} else {
							copy(_egda.P[_ede+2:], _egda.P[_ede+1:])
							_egda.P[_ede+1] = _fdcg
						}
						break
					}
				}
				for _, _ccba := range _egda.Tbl {
					for _, _eea := range _ccba.EG_ContentRowContent {
						for _, _cbcc := range _eea.Tr {
							for _, _aedc := range _cbcc.EG_ContentCellContent {
								for _, _edda := range _aedc.Tc {
									for _, _ccac := range _edda.EG_BlockLevelElts {
										for _, _eaaf := range _ccac.EG_ContentBlockContent {
											for _fcfe, _beafc := range _eaaf.P {
												if _beafc == _bbec {
													_bbgd := _ecgb.X()
													_eaaf.P = append(_eaaf.P, nil)
													if _dcbg {
														copy(_eaaf.P[_fcfe+1:], _eaaf.P[_fcfe:])
														_eaaf.P[_fcfe] = _bbgd
													} else {
														copy(_eaaf.P[_fcfe+2:], _eaaf.P[_fcfe+1:])
														_eaaf.P[_fcfe+1] = _bbgd
													}
													break
												}
											}
										}
									}
								}
							}
						}
					}
				}
				if _egda.Sdt != nil && _egda.Sdt.SdtContent != nil && _egda.Sdt.SdtContent.P != nil {
					for _bge, _dcbe := range _egda.Sdt.SdtContent.P {
						if _dcbe == _bbec {
							_abd := _ecgb.X()
							_egda.Sdt.SdtContent.P = append(_egda.Sdt.SdtContent.P, nil)
							if _dcbg {
								copy(_egda.Sdt.SdtContent.P[_bge+1:], _egda.Sdt.SdtContent.P[_bge:])
								_egda.Sdt.SdtContent.P[_bge] = _abd
							} else {
								copy(_egda.Sdt.SdtContent.P[_bge+2:], _egda.Sdt.SdtContent.P[_bge+1:])
								_egda.Sdt.SdtContent.P[_bge+1] = _abd
							}
							break
						}
					}
				}
			}
		}
	} else {
		_adfe.P = append(_adfe.P, _ecgb.X())
	}
	_badg := _ecgb.Properties()
	if _ddcd, _aga := _badg.Section(); _aga {
		var (
			_gfb  map[string]string
			_gfac map[string]string
		)
		_egaa := _ddcd.X().EG_HdrFtrReferences
		for _, _ddadf := range _egaa {
			if _ddadf.HeaderReference != nil {
				_gfb = map[string]string{_ddadf.HeaderReference.IdAttr: _ddcd._afafb._dab.GetTargetByRelId(_ddadf.HeaderReference.IdAttr)}
			}
			if _ddadf.FooterReference != nil {
				_gfac = map[string]string{_ddadf.FooterReference.IdAttr: _ddcd._afafb._dab.GetTargetByRelId(_ddadf.FooterReference.IdAttr)}
			}
		}
		var _dbcac map[int]common.ImageRef
		for _, _gada := range _ddcd._afafb.Headers() {
			for _bcde, _fbc := range _gfb {
				_afab := fmt.Sprintf("\u0068\u0065\u0061d\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", (_gada.Index() + 1))
				if _afab == _fbc {
					_gfeff := fmt.Sprintf("\u0068\u0065\u0061d\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", _gada.Index())
					_eeb._geb = append(_eeb._geb, _gada.X())
					_cege := _eeb._dab.AddRelationship(_gfeff, unioffice.HeaderType)
					_cege.SetID(_bcde)
					_eeb.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072\u0064\u002f"+_gfeff, "\u0061p\u0070l\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0076\u006e\u0064.\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c\u002e\u0068\u0065\u0061\u0064e\u0072\u002b\u0078\u006d\u006c")
					_eeb._cbfd = append(_eeb._cbfd, common.NewRelationships())
					_fbfe := _gada.Paragraphs()
					for _, _ggdg := range _fbfe {
						for _, _cfbd := range _ggdg.Runs() {
							_fbfa := _cfbd.DrawingAnchored()
							for _, _cagf := range _fbfa {
								if _bbf, _efdf := _cagf.GetImage(); _efdf {
									_dbcac = map[int]common.ImageRef{_gada.Index(): _bbf}
								}
							}
							_agca := _cfbd.DrawingInline()
							for _, _dacb := range _agca {
								if _edge, _abac := _dacb.GetImage(); _abac {
									_dbcac = map[int]common.ImageRef{_gada.Index(): _edge}
								}
							}
						}
					}
				}
			}
		}
		for _fbdca, _efge := range _dbcac {
			for _, _dfbg := range _eeb.Headers() {
				if (_dfbg.Index() + 1) == _fbdca {
					_abda, _ffd := common.ImageFromFile(_efge.Path())
					if _ffd != nil {
						logger.Log.Debug("\u0075\u006e\u0061\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0063r\u0065\u0061\u0074\u0065\u0020\u0069\u006d\u0061\u0067\u0065:\u0020\u0025\u0073", _ffd)
					}
					if _, _ffd = _dfbg.AddImage(_abda); _ffd != nil {
						logger.Log.Debug("u\u006e\u0061\u0062\u006c\u0065\u0020t\u006f\u0020\u0061\u0064\u0064\u0020i\u006d\u0061\u0067\u0065\u0020\u0074\u006f \u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u003a\u0020%\u0073", _ffd)
					}
				}
				for _, _fdcf := range _dfbg.Paragraphs() {
					if _ddcg, _gcd := _ddcd._afafb.Styles.SearchStyleById(_fdcf.Style()); _gcd {
						if _, _gbcbf := _eeb.Styles.SearchStyleById(_fdcf.Style()); !_gbcbf {
							_eeb.Styles.InsertStyle(_ddcg)
						}
					}
				}
			}
		}
		var _eaad map[int]common.ImageRef
		for _, _ggaa := range _ddcd._afafb.Footers() {
			for _edeb, _fbfg := range _gfac {
				_gff := fmt.Sprintf("\u0066\u006f\u006ft\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", (_ggaa.Index() + 1))
				if _gff == _fbfg {
					_bdge := fmt.Sprintf("\u0066\u006f\u006ft\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", _ggaa.Index())
					_eeb._aba = append(_eeb._aba, _ggaa.X())
					_gae := _eeb._dab.AddRelationship(_bdge, unioffice.FooterType)
					_gae.SetID(_edeb)
					_eeb.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072\u0064\u002f"+_bdge, "\u0061p\u0070l\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0076\u006e\u0064.\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c\u002e\u0066\u006f\u006f\u0074e\u0072\u002b\u0078\u006d\u006c")
					_eeb._fdf = append(_eeb._fdf, common.NewRelationships())
					_dbcd := _ggaa.Paragraphs()
					for _, _faac := range _dbcd {
						for _, _dbf := range _faac.Runs() {
							_cafd := _dbf.DrawingAnchored()
							for _, _gcag := range _cafd {
								if _bafe, _begee := _gcag.GetImage(); _begee {
									_eaad = map[int]common.ImageRef{_ggaa.Index(): _bafe}
								}
							}
							_fgcb := _dbf.DrawingInline()
							for _, _gabe := range _fgcb {
								if _bdgc, _ffbg := _gabe.GetImage(); _ffbg {
									_eaad = map[int]common.ImageRef{_ggaa.Index(): _bdgc}
								}
							}
						}
					}
				}
			}
		}
		for _gdb, _fab := range _eaad {
			for _, _debd := range _eeb.Footers() {
				if (_debd.Index() + 1) == _gdb {
					_fbg, _gccb := common.ImageFromFile(_fab.Path())
					if _gccb != nil {
						logger.Log.Debug("\u0075\u006e\u0061\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0063r\u0065\u0061\u0074\u0065\u0020\u0069\u006d\u0061\u0067\u0065:\u0020\u0025\u0073", _gccb)
					}
					if _, _gccb = _debd.AddImage(_fbg); _gccb != nil {
						logger.Log.Debug("u\u006e\u0061\u0062\u006c\u0065\u0020t\u006f\u0020\u0061\u0064\u0064\u0020i\u006d\u0061\u0067\u0065\u0020\u0074\u006f \u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u003a\u0020%\u0073", _gccb)
					}
				}
				for _, _eee := range _debd.Paragraphs() {
					if _bgfe, _aedf := _ddcd._afafb.Styles.SearchStyleById(_eee.Style()); _aedf {
						if _, _aaada := _eeb.Styles.SearchStyleById(_eee.Style()); !_aaada {
							_eeb.Styles.InsertStyle(_bgfe)
						}
					}
				}
			}
		}
	}
	_dcf := _ecgb.Numbering()
	_eeb.Numbering._cbag.AbstractNum = append(_eeb.Numbering._cbag.AbstractNum, _dcf._cbag.AbstractNum...)
	_eeb.Numbering._cbag.Num = append(_eeb.Numbering._cbag.Num, _dcf._cbag.Num...)
	return Paragraph{_eeb, _ecgb.X()}
}

// AddDefinition adds a new numbering definition.
func (_aadf Numbering) AddDefinition() NumberingDefinition {
	_cefbg := wml.NewCT_Num()
	_gfgea := int64(1)
	for _, _deff := range _aadf.Definitions() {
		if _deff.AbstractNumberID() >= _gfgea {
			_gfgea = _deff.AbstractNumberID() + 1
		}
	}
	_defg := int64(1)
	for _, _faag := range _aadf.X().Num {
		if _faag.NumIdAttr >= _defg {
			_defg = _faag.NumIdAttr + 1
		}
	}
	_cefbg.NumIdAttr = _defg
	_cefbg.AbstractNumId = wml.NewCT_DecimalNumber()
	_cefbg.AbstractNumId.ValAttr = _gfgea
	_adcg := wml.NewCT_AbstractNum()
	_adcg.AbstractNumIdAttr = _gfgea
	_aadf._cbag.AbstractNum = append(_aadf._cbag.AbstractNum, _adcg)
	_aadf._cbag.Num = append(_aadf._cbag.Num, _cefbg)
	return NumberingDefinition{_adcg}
}

// AddParagraph adds a paragraph to the footnote.
func (_abef Footnote) AddParagraph() Paragraph {
	_fcdf := wml.NewEG_ContentBlockContent()
	_acbe := len(_abef._bgcda.EG_BlockLevelElts[0].EG_ContentBlockContent)
	_abef._bgcda.EG_BlockLevelElts[0].EG_ContentBlockContent = append(_abef._bgcda.EG_BlockLevelElts[0].EG_ContentBlockContent, _fcdf)
	_ffdf := wml.NewCT_P()
	var _gcba *wml.CT_String
	if _acbe != 0 {
		_ebeg := len(_abef._bgcda.EG_BlockLevelElts[0].EG_ContentBlockContent[_acbe-1].P)
		_gcba = _abef._bgcda.EG_BlockLevelElts[0].EG_ContentBlockContent[_acbe-1].P[_ebeg-1].PPr.PStyle
	} else {
		_gcba = wml.NewCT_String()
		_gcba.ValAttr = "\u0046\u006f\u006f\u0074\u006e\u006f\u0074\u0065"
	}
	_fcdf.P = append(_fcdf.P, _ffdf)
	_bcda := Paragraph{_abef._gffg, _ffdf}
	_bcda._eagd.PPr = wml.NewCT_PPr()
	_bcda._eagd.PPr.PStyle = _gcba
	_bcda._eagd.PPr.RPr = wml.NewCT_ParaRPr()
	return _bcda
}

// ItalicValue returns the precise nature of the italic setting (unset, off or on).
func (_bbbf RunProperties) ItalicValue() OnOffValue { return _fgccb(_bbbf._gbdb.I) }

// AddWatermarkText adds new watermark text to the document.
func (_fbfc *Document) AddWatermarkText(text string) WatermarkText {
	var _eaf []Header
	if _ecc, _cfef := _fbfc.BodySection().GetHeader(wml.ST_HdrFtrDefault); _cfef {
		_eaf = append(_eaf, _ecc)
	}
	if _fcd, _bcf := _fbfc.BodySection().GetHeader(wml.ST_HdrFtrEven); _bcf {
		_eaf = append(_eaf, _fcd)
	}
	if _bcce, _ffcd := _fbfc.BodySection().GetHeader(wml.ST_HdrFtrFirst); _ffcd {
		_eaf = append(_eaf, _bcce)
	}
	if len(_eaf) < 1 {
		_cgaff := _fbfc.AddHeader()
		_fbfc.BodySection().SetHeader(_cgaff, wml.ST_HdrFtrDefault)
		_eaf = append(_eaf, _cgaff)
	}
	_eabg := NewWatermarkText()
	for _, _gcbf := range _eaf {
		_dfggg := _gcbf.Paragraphs()
		if len(_dfggg) < 1 {
			_bgfa := _gcbf.AddParagraph()
			_bgfa.AddRun().AddText("")
		}
		for _, _gea := range _gcbf.X().EG_ContentBlockContent {
			for _, _dcg := range _gea.P {
				for _, _edde := range _dcg.EG_PContent {
					for _, _gdaa := range _edde.EG_ContentRunContent {
						if _gdaa.R == nil {
							continue
						}
						for _, _dgae := range _gdaa.R.EG_RunInnerContent {
							_dgae.Pict = _eabg._cegfa
							break
						}
					}
				}
			}
		}
	}
	_eabg.SetText(text)
	return _eabg
}

// SetAll sets all of the borders to a given value.
func (_eagb TableBorders) SetAll(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_eagb.SetBottom(t, c, thickness)
	_eagb.SetLeft(t, c, thickness)
	_eagb.SetRight(t, c, thickness)
	_eagb.SetTop(t, c, thickness)
	_eagb.SetInsideHorizontal(t, c, thickness)
	_eagb.SetInsideVertical(t, c, thickness)
}

// Nodes contains slice of Node element.
type Nodes struct{ _gabfc []Node }

func _cbcf(_ccbg *Document, _ccbae Paragraph) listItemInfo {
	if _ccbg.Numbering.X() == nil {
		return listItemInfo{}
	}
	if len(_ccbg.Numbering.Definitions()) < 1 {
		return listItemInfo{}
	}
	_efce := _bgcg(_ccbae)
	if _efce == nil {
		return listItemInfo{}
	}
	_gbafb := _ccbg.GetNumberingLevelByIds(_efce.NumId.ValAttr, _efce.Ilvl.ValAttr)
	if _ffcdc := _gbafb.X(); _ffcdc == nil {
		return listItemInfo{}
	}
	_aggb := int64(0)
	for _, _bgca := range _ccbg.Numbering._cbag.Num {
		if _bgca != nil && _bgca.NumIdAttr == _efce.NumId.ValAttr {
			_aggb = _bgca.AbstractNumId.ValAttr
		}
	}
	return listItemInfo{FromParagraph: &_ccbae, AbstractNumId: &_aggb, NumberingLevel: &_gbafb}
}
func (_ace *chart) X() *dmlChart.ChartSpace { return _ace._ffb }
func (_eedbd *WatermarkPicture) getShape() *unioffice.XSDAny {
	return _eedbd.getInnerElement("\u0073\u0068\u0061p\u0065")
}

// WatermarkText is watermark text within the document.
type WatermarkText struct {
	_cegfa *wml.CT_Picture
	_edfdc *vmldrawing.TextpathStyle
	_bfbf  *vml.Shape
	_gafdb *vml.Shapetype
}

// ParagraphSpacing controls the spacing for a paragraph and its lines.
type ParagraphSpacing struct{ _ffede *wml.CT_Spacing }

// X returns the inner wrapped XML type.
func (_eedc Numbering) X() *wml.Numbering { return _eedc._cbag }
func (_accb Footnote) id() int64          { return _accb._bgcda.IdAttr }

// SetPictureSize set watermark picture size with given width and height.
func (_gcafe *WatermarkPicture) SetPictureSize(width, height int64) {
	if _gcafe._fdgfa != nil {
		_ggffe := _gcafe.GetShapeStyle()
		_ggffe.SetWidth(float64(width) * measurement.Point)
		_ggffe.SetHeight(float64(height) * measurement.Point)
		_gcafe.SetShapeStyle(_ggffe)
	}
}

// AddTabStop adds a tab stop to the paragraph.  It controls the position of text when using Run.AddTab()
func (_gefd ParagraphProperties) AddTabStop(position measurement.Distance, justificaton wml.ST_TabJc, leader wml.ST_TabTlc) {
	if _gefd._dfaf.Tabs == nil {
		_gefd._dfaf.Tabs = wml.NewCT_Tabs()
	}
	_fceddc := wml.NewCT_TabStop()
	_fceddc.LeaderAttr = leader
	_fceddc.ValAttr = justificaton
	_fceddc.PosAttr.Int64 = unioffice.Int64(int64(position / measurement.Twips))
	_gefd._dfaf.Tabs.Tab = append(_gefd._dfaf.Tabs.Tab, _fceddc)
}

// RemoveEndnote removes a endnote from both the paragraph and the document
// the requested endnote must be anchored on the paragraph being referenced.
func (_ffba Paragraph) RemoveEndnote(id int64) {
	_ebefb := _ffba._fagf._ccb
	var _dcee int
	for _fggc, _dcbfc := range _ebefb.CT_Endnotes.Endnote {
		if _dcbfc.IdAttr == id {
			_dcee = _fggc
		}
	}
	_dcee = 0
	_ebefb.CT_Endnotes.Endnote[_dcee] = nil
	_ebefb.CT_Endnotes.Endnote[_dcee] = _ebefb.CT_Endnotes.Endnote[len(_ebefb.CT_Endnotes.Endnote)-1]
	_ebefb.CT_Endnotes.Endnote = _ebefb.CT_Endnotes.Endnote[:len(_ebefb.CT_Endnotes.Endnote)-1]
	var _fged Run
	for _, _ecfa := range _ffba.Runs() {
		if _dgfgb, _faaf := _ecfa.IsEndnote(); _dgfgb {
			if _faaf == id {
				_fged = _ecfa
			}
		}
	}
	_ffba.RemoveRun(_fged)
}

// SetOutlineLevel sets the outline level of this style.
func (_fecf ParagraphStyleProperties) SetOutlineLevel(lvl int) {
	_fecf._gfee.OutlineLvl = wml.NewCT_DecimalNumber()
	_fecf._gfee.OutlineLvl.ValAttr = int64(lvl)
}
func (_dedg *Document) validateTableCells() error {
	for _, _gbgc := range _dedg.doc.Body.EG_BlockLevelElts {
		for _, _baacb := range _gbgc.EG_ContentBlockContent {
			for _, _edf := range _baacb.Tbl {
				for _, _efff := range _edf.EG_ContentRowContent {
					for _, _agef := range _efff.Tr {
						_ggdb := false
						for _, _cbga := range _agef.EG_ContentCellContent {
							_dadb := false
							for _, _eade := range _cbga.Tc {
								_ggdb = true
								for _, _caef := range _eade.EG_BlockLevelElts {
									for _, _dfd := range _caef.EG_ContentBlockContent {
										if len(_dfd.P) > 0 {
											_dadb = true
											break
										}
									}
								}
							}
							if !_dadb {
								return errors.New("t\u0061\u0062\u006c\u0065\u0020\u0063e\u006c\u006c\u0020\u006d\u0075\u0073t\u0020\u0063\u006f\u006e\u0074\u0061\u0069n\u0020\u0061\u0020\u0070\u0061\u0072\u0061\u0067\u0072\u0061p\u0068")
							}
						}
						if !_ggdb {
							return errors.New("\u0074\u0061b\u006c\u0065\u0020\u0072\u006f\u0077\u0020\u006d\u0075\u0073\u0074\u0020\u0063\u006f\u006e\u0074\u0061\u0069\u006e\u0020\u0061\u0020ce\u006c\u006c")
						}
					}
				}
			}
		}
	}
	return nil
}

// SetWidth sets the table with to a specified width.
func (_ddfa TableProperties) SetWidth(d measurement.Distance) {
	_ddfa._efag.TblW = wml.NewCT_TblWidth()
	_ddfa._efag.TblW.TypeAttr = wml.ST_TblWidthDxa
	_ddfa._efag.TblW.WAttr = &wml.ST_MeasurementOrPercent{}
	_ddfa._efag.TblW.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_ddfa._efag.TblW.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(d / measurement.Twips))
}

// SetEastAsiaTheme sets the font East Asia Theme.
func (_abe Fonts) SetEastAsiaTheme(t wml.ST_Theme) { _abe._feae.EastAsiaThemeAttr = t }

// AppendNode append node to document element.
func (_gbfe *Document) AppendNode(node Node) {
	_gbfe.insertImageFromNode(node)
	_gbfe.insertStyleFromNode(node)
	for _, _agae := range node.Children {
		_gbfe.insertImageFromNode(_agae)
		_gbfe.insertStyleFromNode(_agae)
	}
	switch _aaeg := node.X().(type) {
	case *Paragraph:
		_gbfe.appendParagraph(nil, *_aaeg, false)
	case *Table:
		_gbfe.appendTable(nil, *_aaeg, false)
	}
	if node._cdbd != nil {
		if node._cdbd._ffbc != nil {
			if _bdbf := _gbfe._dab.FindRIDForN(0, unioffice.ThemeType); _bdbf == "" {
				if _bgdgd := node._cdbd._dab.FindRIDForN(0, unioffice.ThemeType); _bgdgd != "" {
					_gbfe._ffbc = append(_gbfe._ffbc, node._cdbd._ffbc...)
					_dbe := node._cdbd._dab.GetTargetByRelId(_bgdgd)
					_gbfe.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072\u0064\u002f"+_dbe, "\u0061\u0070\u0070\u006c\u0069\u0063\u0061t\u0069\u006f\u006e/\u0076\u006e\u0064.\u006f\u0070e\u006e\u0078\u006d\u006c\u0066\u006fr\u006dat\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0074\u0068\u0065\u006d\u0065\u002b\u0078\u006d\u006c")
					_gbfe._dab.AddRelationship(_dbe, unioffice.ThemeType)
				}
			}
		}
		_bgag := _gbfe._eaa
		_cecae := node._cdbd._eaa
		if _bgag != nil {
			if _cecae != nil {
				if _bgag.Font != nil {
					if _cecae.Font != nil {
						for _, _cfaaf := range _cecae.Font {
							_daea := true
							for _, _gdfbe := range _bgag.Font {
								if _gdfbe.NameAttr == _cfaaf.NameAttr {
									_daea = false
									break
								}
							}
							if _daea {
								_bgag.Font = append(_bgag.Font, _cfaaf)
							}
						}
					}
				} else {
					_bgag.Font = _cecae.Font
				}
			}
		} else if _cecae != nil {
			_bgag = _cecae
		}
		_gbfe._eaa = _bgag
		if _gfdg := _gbfe._dab.FindRIDForN(0, unioffice.FontTableType); _gfdg == "" {
			_gbfe.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072d/\u0066\u006f\u006e\u0074\u0054\u0061\u0062\u006c\u0065\u002e\u0078\u006d\u006c", "\u0061\u0070\u0070\u006c\u0069c\u0061\u0074\u0069\u006f\u006e\u002f\u0076n\u0064\u002e\u006f\u0070\u0065\u006e\u0078\u006d\u006c\u0066\u006f\u0072\u006d\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063e\u0073\u0073\u0069\u006e\u0067\u006d\u006c\u002e\u0066\u006f\u006e\u0074T\u0061\u0062\u006c\u0065\u002b\u0078m\u006c")
			_gbfe._dab.AddRelationship("\u0066\u006f\u006e\u0074\u0054\u0061\u0062\u006c\u0065\u002e\u0078\u006d\u006c", unioffice.FontTableType)
		}
	}
}

// AddFieldWithFormatting adds a field (automatically computed text) to the
// document with field specifc formatting.
func (_dfad Run) AddFieldWithFormatting(code string, fmt string, isDirty bool) {
	_fdfb := _dfad.newIC()
	_fdfb.FldChar = wml.NewCT_FldChar()
	_fdfb.FldChar.FldCharTypeAttr = wml.ST_FldCharTypeBegin
	if isDirty {
		_fdfb.FldChar.DirtyAttr = &sharedTypes.ST_OnOff{}
		_fdfb.FldChar.DirtyAttr.Bool = unioffice.Bool(true)
	}
	_fdfb = _dfad.newIC()
	_fdfb.InstrText = wml.NewCT_Text()
	if fmt != "" {
		_fdfb.InstrText.Content = code + "\u0020" + fmt
	} else {
		_fdfb.InstrText.Content = code
	}
	_fdfb = _dfad.newIC()
	_fdfb.FldChar = wml.NewCT_FldChar()
	_fdfb.FldChar.FldCharTypeAttr = wml.ST_FldCharTypeEnd
}

// Cells returns the cells defined in the table.
func (_cffcb Row) Cells() []Cell {
	_gcfe := []Cell{}
	for _, _bgcac := range _cffcb.ctRow.EG_ContentCellContent {
		for _, _gdbge := range _bgcac.Tc {
			_gcfe = append(_gcfe, Cell{_cffcb.doc, _gdbge})
		}
		if _bgcac.Sdt != nil && _bgcac.Sdt.SdtContent != nil {
			for _, _ddef := range _bgcac.Sdt.SdtContent.Tc {
				_gcfe = append(_gcfe, Cell{_cffcb.doc, _ddef})
			}
		}
	}
	return _gcfe
}

// Definitions returns the defined numbering definitions.
func (_egged Numbering) Definitions() []NumberingDefinition {
	_cbffa := []NumberingDefinition{}
	if _egged._cbag != nil {
		for _, _aefa := range _egged._cbag.AbstractNum {
			_cbffa = append(_cbffa, NumberingDefinition{_aefa})
		}
	}
	return _cbffa
}

// SetBefore sets the spacing that comes before the paragraph.
func (_gcdae ParagraphSpacing) SetBefore(before measurement.Distance) {
	_gcdae._ffede.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
	_gcdae._ffede.BeforeAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(before / measurement.Twips))
}

// X returns the inner wrapped XML type.
func (_eeda TableProperties) X() *wml.CT_TblPr { return _eeda._efag }

// Close closes the document, removing any temporary files that might have been
// created when opening a document.
func (_beafg *Document) Close() error {
	if _beafg.TmpPath != "" {
		return tempstorage.RemoveAll(_beafg.TmpPath)
	}
	return nil
}

// TableConditionalFormatting returns a conditional formatting object of a given
// type.  Calling this method repeatedly will return the same object.
func (_bbged Style) TableConditionalFormatting(typ wml.ST_TblStyleOverrideType) TableConditionalFormatting {
	for _, _gcbe := range _bbged._gaege.TblStylePr {
		if _gcbe.TypeAttr == typ {
			return TableConditionalFormatting{_gcbe}
		}
	}
	_caded := wml.NewCT_TblStylePr()
	_caded.TypeAttr = typ
	_bbged._gaege.TblStylePr = append(_bbged._gaege.TblStylePr, _caded)
	return TableConditionalFormatting{_caded}
}

// SetSemiHidden controls if the style is hidden in the UI.
func (_gagaa Style) SetSemiHidden(b bool) {
	if b {
		_gagaa._gaege.SemiHidden = wml.NewCT_OnOff()
	} else {
		_gagaa._gaege.SemiHidden = nil
	}
}

// GetImageObjByRelId returns a common.Image with the associated relation ID in the
// document.
func (_dadcc *Document) GetImageObjByRelId(relId string) (common.Image, error) {
	_fbff := _dadcc._dab.GetTargetByRelId(relId)
	return _dadcc.DocBase.GetImageBytesByTarget(_fbff)
}

// SetInsideVertical sets the interior vertical borders to a specified type, color and thickness.
func (_dfb CellBorders) SetInsideVertical(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_dfb._gf.InsideV = wml.NewCT_Border()
	_feadc(_dfb._gf.InsideV, t, c, thickness)
}

// SetRowBandSize sets the number of Rows in the row band
func (_cfbddf TableStyleProperties) SetRowBandSize(rows int64) {
	_cfbddf._degc.TblStyleRowBandSize = wml.NewCT_DecimalNumber()
	_cfbddf._degc.TblStyleRowBandSize.ValAttr = rows
}

// Borders allows manipulation of the table borders.
func (_dbfffd TableProperties) Borders() TableBorders {
	if _dbfffd._efag.TblBorders == nil {
		_dbfffd._efag.TblBorders = wml.NewCT_TblBorders()
	}
	return TableBorders{_dbfffd._efag.TblBorders}
}
func _dfedg() *vml.Path {
	_efbc := vml.NewPath()
	_efbc.ExtrusionokAttr = sharedTypes.ST_TrueFalseTrue
	_efbc.GradientshapeokAttr = sharedTypes.ST_TrueFalseTrue
	_efbc.ConnecttypeAttr = vml.OfcST_ConnectTypeRect
	return _efbc
}
func (_caee *Document) InsertTableBefore(relativeTo Paragraph) Table {
	return _caee.insertTable(relativeTo, true)
}

// SetDefaultValue sets the default value of a FormFieldTypeDropDown. For
// FormFieldTypeDropDown, the value must be one of the fields possible values.
func (_afgd FormField) SetDefaultValue(v string) {
	if _afgd._cbde.DdList != nil {
		for _bfea, _ebec := range _afgd.PossibleValues() {
			if _ebec == v {
				_afgd._cbde.DdList.Default = wml.NewCT_DecimalNumber()
				_afgd._cbde.DdList.Default.ValAttr = int64(_bfea)
				break
			}
		}
	}
}

// AddTextInput adds text input form field to the paragraph and returns it.
func (_dgfb Paragraph) AddTextInput(name string) FormField {
	_fddd := _dgfb.addFldCharsForField(name, "\u0046\u004f\u0052\u004d\u0054\u0045\u0058\u0054")
	_fddd._cbde.TextInput = wml.NewCT_FFTextInput()
	return _fddd
}
func (_bcg *chart) RelId() string { return _bcg._fda }

// SetLineSpacing controls the line spacing of the paragraph.
func (_cccg ParagraphStyleProperties) SetLineSpacing(m measurement.Distance, rule wml.ST_LineSpacingRule) {
	if _cccg._gfee.Spacing == nil {
		_cccg._gfee.Spacing = wml.NewCT_Spacing()
	}
	if rule == wml.ST_LineSpacingRuleUnset {
		_cccg._gfee.Spacing.LineRuleAttr = wml.ST_LineSpacingRuleUnset
		_cccg._gfee.Spacing.LineAttr = nil
	} else {
		_cccg._gfee.Spacing.LineRuleAttr = rule
		_cccg._gfee.Spacing.LineAttr = &wml.ST_SignedTwipsMeasure{}
		_cccg._gfee.Spacing.LineAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// SetCSTheme sets the font complex script theme.
func (_eaca Fonts) SetCSTheme(t wml.ST_Theme) { _eaca._feae.CsthemeAttr = t }
func (_egad *WatermarkText) getShapeType() *unioffice.XSDAny {
	return _egad.getInnerElement("\u0073h\u0061\u0070\u0065\u0074\u0079\u0070e")
}

// SetTop sets the top border to a specified type, color and thickness.
func (_bee CellBorders) SetTop(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_bee._gf.Top = wml.NewCT_Border()
	_feadc(_bee._gf.Top, t, c, thickness)
}

// SetTextWrapTight sets the text wrap to tight with a give wrap type.
func (_fgc AnchoredDrawing) SetTextWrapTight(option *AnchorDrawWrapOptions) {
	_fgc._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_fgc._dgc.Choice.WrapTight = wml.NewWdCT_WrapTight()
	_fgc._dgc.Choice.WrapTight.WrapTextAttr = wml.WdST_WrapTextBothSides
	_aae := false
	_fgc._dgc.Choice.WrapTight.WrapPolygon.EditedAttr = &_aae
	if option == nil {
		option = NewAnchorDrawWrapOptions()
	}
	_fgc._dgc.Choice.WrapTight.WrapPolygon.LineTo = option.GetWrapPathLineTo()
	_fgc._dgc.Choice.WrapTight.WrapPolygon.Start = option.GetWrapPathStart()
	_fgc._dgc.LayoutInCellAttr = true
	_fgc._dgc.AllowOverlapAttr = true
}

// SetStyle sets the style of a paragraph and is identical to setting it on the
// paragraph's Properties()
func (_acgca Paragraph) SetStyle(s string) {
	_acgca.ensurePPr()
	if s == "" {
		_acgca._eagd.PPr.PStyle = nil
	} else {
		_acgca._eagd.PPr.PStyle = wml.NewCT_String()
		_acgca._eagd.PPr.PStyle.ValAttr = s
	}
}

// EastAsiaFont returns the name of paragraph font family for East Asia.
func (_agbe ParagraphProperties) EastAsiaFont() string {
	if _efgdg := _agbe._dfaf.RPr.RFonts; _efgdg != nil {
		if _efgdg.EastAsiaAttr != nil {
			return *_efgdg.EastAsiaAttr
		}
	}
	return ""
}

// DoubleStrike returns true if paragraph is double striked.
func (_afec ParagraphProperties) DoubleStrike() bool { return _cadf(_afec._dfaf.RPr.Dstrike) }

// ClearColor clears the text color.
func (_abadc RunProperties) ClearColor() { _abadc._gbdb.Color = nil }

// SearchStyleByName return style by its name.
func (_gbgcad Styles) SearchStyleByName(name string) (Style, bool) {
	for _, _eeccb := range _gbgcad._abca.Style {
		if _eeccb.Name != nil {
			if _eeccb.Name.ValAttr == name {
				return Style{_eeccb}, true
			}
		}
	}
	return Style{}, false
}

// SetHangingIndent controls special indent of paragraph.
func (_eaeg Paragraph) SetHangingIndent(m measurement.Distance) {
	_eaeg.ensurePPr()
	_dbdd := _eaeg._eagd.PPr
	if _dbdd.Ind == nil {
		_dbdd.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_dbdd.Ind.HangingAttr = nil
	} else {
		_dbdd.Ind.HangingAttr = &sharedTypes.ST_TwipsMeasure{}
		_dbdd.Ind.HangingAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

// Clear clears all content within a header
func (_gefg Header) Clear() { _gefg._deae.EG_ContentBlockContent = nil }

// SetAlignment positions an anchored image via alignment.  Offset is
// incompatible with SetOffset, whichever is called last is applied.
func (_cfd AnchoredDrawing) SetAlignment(h wml.WdST_AlignH, v wml.WdST_AlignV) {
	_cfd.SetHAlignment(h)
	_cfd.SetVAlignment(v)
}

// X returns the inner wrapped XML type.
func (_da AnchoredDrawing) X() *wml.WdAnchor { return _da._dgc }

// SetAfterAuto controls if spacing after a paragraph is automatically determined.
func (_gdfc ParagraphSpacing) SetAfterAuto(b bool) {
	if b {
		_gdfc._ffede.AfterAutospacingAttr = &sharedTypes.ST_OnOff{}
		_gdfc._ffede.AfterAutospacingAttr.Bool = unioffice.Bool(true)
	} else {
		_gdfc._ffede.AfterAutospacingAttr = nil
	}
}

// X returns the inner wrapped XML type.
func (_dcgbe RunProperties) X() *wml.CT_RPr { return _dcgbe._gbdb }

// Strike returns true if run is striked.
func (_aecd RunProperties) Strike() bool { return _cadf(_aecd._gbdb.Strike) }

// SetFollowImageShape sets wrapPath to follow image shape,
// if nil return wrapPath that follow image size.
func (_egg AnchorDrawWrapOptions) SetFollowImageShape(val bool) {
	_egg._cef = val
	if !val {
		_ebg, _fgec := _afa()
		_egg._dd = _ebg
		_egg._cbf = _fgec
	}
}

// BodySection returns the default body section used for all preceding
// paragraphs until the previous Section. If there is no previous sections, the
// body section applies to the entire document.
func (_bbc *Document) BodySection() Section {
	if _bbc.doc.Body.SectPr == nil {
		_bbc.doc.Body.SectPr = wml.NewCT_SectPr()
	}
	return Section{_bbc, _bbc.doc.Body.SectPr}
}

// X returns the inner wrapped XML type.
func (_afff Color) X() *wml.CT_Color { return _afff._ec }

// SetTextWrapThrough sets the text wrap to through with a give wrap type.
func (_ge AnchoredDrawing) SetTextWrapThrough(option *AnchorDrawWrapOptions) {
	_ge._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_ge._dgc.Choice.WrapThrough = wml.NewWdCT_WrapThrough()
	_ge._dgc.Choice.WrapThrough.WrapTextAttr = wml.WdST_WrapTextBothSides
	_fb := false
	_ge._dgc.Choice.WrapThrough.WrapPolygon.EditedAttr = &_fb
	if option == nil {
		option = NewAnchorDrawWrapOptions()
	}
	_ge._dgc.Choice.WrapThrough.WrapPolygon.Start = option.GetWrapPathStart()
	_ge._dgc.Choice.WrapThrough.WrapPolygon.LineTo = option.GetWrapPathLineTo()
	_ge._dgc.LayoutInCellAttr = true
	_ge._dgc.AllowOverlapAttr = true
}

// SetCellSpacingPercent sets the cell spacing within a table to a percent width.
func (_afbaf TableProperties) SetCellSpacingPercent(pct float64) {
	_afbaf._efag.TblCellSpacing = wml.NewCT_TblWidth()
	_afbaf._efag.TblCellSpacing.TypeAttr = wml.ST_TblWidthPct
	_afbaf._efag.TblCellSpacing.WAttr = &wml.ST_MeasurementOrPercent{}
	_afbaf._efag.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_afbaf._efag.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(pct * 50))
}

// SetBottomPct sets the cell bottom margin
func (_geg CellMargins) SetBottomPct(pct float64) {
	_geg._cdae.Bottom = wml.NewCT_TblWidth()
	_aff(_geg._cdae.Bottom, pct)
}

// X returns the inner wrapped XML type.
func (_cbff CellProperties) X() *wml.CT_TcPr { return _cbff._cgc }

// AddFooter creates a Footer associated with the document, but doesn't add it
// to the document for display.
func (_aca *Document) AddFooter() Footer {
	_dabf := wml.NewFtr()
	_aca._aba = append(_aca._aba, _dabf)
	_abbc := fmt.Sprintf("\u0066\u006f\u006ft\u0065\u0072\u0025\u0064\u002e\u0078\u006d\u006c", len(_aca._aba))
	_aca._dab.AddRelationship(_abbc, unioffice.FooterType)
	_aca.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072\u0064\u002f"+_abbc, "\u0061p\u0070l\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0076\u006e\u0064.\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c\u002e\u0066\u006f\u006f\u0074e\u0072\u002b\u0078\u006d\u006c")
	_aca._fdf = append(_aca._fdf, common.NewRelationships())
	return Footer{_aca, _dabf}
}

// IgnoreSpaceBetweenParagraphOfSameStyle sets contextual spacing.
func (_eebgf Paragraph) IgnoreSpaceBetweenParagraphOfSameStyle() {
	_eebgf.ensurePPr()
	if _eebgf._eagd.PPr.ContextualSpacing == nil {
		_eebgf._eagd.PPr.ContextualSpacing = wml.NewCT_OnOff()
	}
	_eebgf._eagd.PPr.ContextualSpacing.ValAttr = &sharedTypes.ST_OnOff{ST_OnOff1: sharedTypes.ST_OnOff1On}
}

// NewWatermarkText generates a new WatermarkText.
func NewWatermarkText() WatermarkText {
	_eeeae := vml.NewShapetype()
	_ebgee := vml.NewEG_ShapeElements()
	_ebgee.Formulas = _cgcae()
	_ebgee.Path = _cegfac()
	_ebgee.Textpath = _agefe()
	_ebgee.Handles = _bdcf()
	_ebgee.Lock = _gcad()
	_eeeae.EG_ShapeElements = []*vml.EG_ShapeElements{_ebgee}
	var (
		_adfc  = "_\u0078\u0030\u0030\u0030\u0030\u005f\u0074\u0031\u0033\u0036"
		_edfec = "2\u0031\u0036\u0030\u0030\u002c\u0032\u0031\u0036\u0030\u0030"
		_gadee = float32(136.0)
		_ddbf  = "\u0031\u0030\u00380\u0030"
		_dcegf = "m\u0040\u0037\u002c\u006c\u0040\u0038,\u006d\u0040\u0035\u002c\u0032\u0031\u0036\u0030\u0030l\u0040\u0036\u002c2\u00316\u0030\u0030\u0065"
	)
	_eeeae.IdAttr = &_adfc
	_eeeae.CoordsizeAttr = &_edfec
	_eeeae.SptAttr = &_gadee
	_eeeae.AdjAttr = &_ddbf
	_eeeae.PathAttr = &_dcegf
	_aead := vml.NewShape()
	_cefd := vml.NewEG_ShapeElements()
	_cefd.Textpath = _cdbee()
	_aead.EG_ShapeElements = []*vml.EG_ShapeElements{_cefd}
	var (
		_gacbfg = "\u0050\u006f\u0077\u0065\u0072\u0050l\u0075\u0073\u0057\u0061\u0074\u0065\u0072\u004d\u0061\u0072\u006b\u004f\u0062j\u0065\u0063\u0074\u0031\u0033\u0036\u00380\u0030\u0038\u0038\u0036"
		_cfbg   = "\u005f\u0078\u00300\u0030\u0030\u005f\u0073\u0032\u0030\u0035\u0031"
		_ddfaf  = "\u0023\u005f\u00780\u0030\u0030\u0030\u005f\u0074\u0031\u0033\u0036"
		_adecd  = ""
		_aedaa  = "\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e\u003a\u0061\u0062\u0073\u006f\u006c\u0075\u0074\u0065\u003b\u006d\u0061\u0072\u0067\u0069\u006e\u002d\u006c\u0065f\u0074:\u0030\u003b\u006d\u0061\u0072\u0067\u0069\u006e\u002d\u0074o\u0070\u003a\u0030\u003b\u0077\u0069\u0064\u0074\u0068\u003a\u0034\u0036\u0038\u0070\u0074;\u0068\u0065\u0069\u0067\u0068\u0074\u003a\u0032\u0033\u0034\u0070\u0074\u003b\u007a\u002d\u0069\u006ede\u0078\u003a\u002d\u0032\u0035\u0031\u0036\u0035\u0031\u0030\u0037\u0032\u003b\u006d\u0073\u006f\u002d\u0077\u0072\u0061\u0070\u002d\u0065\u0064\u0069\u0074\u0065\u0064\u003a\u0066\u003b\u006d\u0073\u006f\u002d\u0077\u0069\u0064\u0074\u0068\u002d\u0070\u0065\u0072\u0063\u0065\u006e\u0074\u003a\u0030\u003b\u006d\u0073\u006f\u002d\u0068\u0065\u0069\u0067h\u0074-p\u0065\u0072\u0063\u0065\u006et\u003a\u0030\u003b\u006d\u0073\u006f\u002d\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e\u002d\u0068\u006f\u0072\u0069\u007a\u006fn\u0074\u0061\u006c\u003a\u0063\u0065\u006e\u0074\u0065\u0072\u003b\u006d\u0073\u006f\u002d\u0070o\u0073\u0069\u0074\u0069\u006f\u006e\u002d\u0068\u006f\u0072\u0069\u007a\u006f\u006e\u0074\u0061\u006c\u002d\u0072\u0065l\u0061\u0074\u0069\u0076\u0065:\u006d\u0061\u0072\u0067\u0069n\u003b\u006d\u0073o\u002d\u0070\u006f\u0073\u0069\u0074\u0069o\u006e-\u0076\u0065\u0072\u0074\u0069\u0063\u0061\u006c\u003a\u0063\u0065\u006e\u0074\u0065\u0072\u003b\u006d\u0073\u006f\u002d\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e\u002d\u0076\u0065r\u0074\u0069\u0063\u0061\u006c\u002d\u0072e\u006c\u0061\u0074i\u0076\u0065\u003a\u006d\u0061\u0072\u0067\u0069\u006e\u003b\u006d\u0073\u006f\u002d\u0077\u0069\u0064\u0074\u0068\u002d\u0070\u0065\u0072\u0063e\u006e\u0074\u003a\u0030\u003b\u006d\u0073\u006f\u002dh\u0065\u0069\u0067\u0068t\u002dp\u0065\u0072\u0063\u0065\u006et\u003a0"
		_bfac   = "\u0073\u0069\u006c\u0076\u0065\u0072"
	)
	_aead.IdAttr = &_gacbfg
	_aead.SpidAttr = &_cfbg
	_aead.TypeAttr = &_ddfaf
	_aead.AltAttr = &_adecd
	_aead.StyleAttr = &_aedaa
	_aead.AllowincellAttr = sharedTypes.ST_TrueFalseFalse
	_aead.FillcolorAttr = &_bfac
	_aead.StrokedAttr = sharedTypes.ST_TrueFalseFalse
	_aggbc := wml.NewCT_Picture()
	_aggbc.Any = []unioffice.Any{_eeeae, _aead}
	return WatermarkText{_cegfa: _aggbc, _bfbf: _aead, _gafdb: _eeeae}
}

// X returns the inner wrapped XML type.
func (_eeca NumberingLevel) X() *wml.CT_Lvl { return _eeca.lvl }
func (_cafg Document) mergeFields() []mergeFieldInfo {
	_bbcc := []Paragraph{}
	_cgff := []mergeFieldInfo{}
	for _, _bccd := range _cafg.Tables() {
		for _, _cece := range _bccd.Rows() {
			for _, _aeae := range _cece.Cells() {
				_bbcc = append(_bbcc, _aeae.Paragraphs()...)
			}
		}
	}
	_bbcc = append(_bbcc, _cafg.Paragraphs()...)
	for _, _bcfe := range _bbcc {
		_dggb := _bcfe.Runs()
		_dgac := -1
		_edca := -1
		_bbde := -1
		_dffb := mergeFieldInfo{}
		for _, _dbaad := range _bcfe._eagd.EG_PContent {
			for _, _bgdb := range _dbaad.FldSimple {
				if strings.Contains(_bgdb.InstrAttr, "\u004d\u0045\u0052\u0047\u0045\u0046\u0049\u0045\u004c\u0044") {
					_gcbcc := _cadg(_bgdb.InstrAttr)
					_gcbcc._debdg = true
					_gcbcc._abdbd = _bcfe
					_gcbcc._ceaf = _dbaad
					_cgff = append(_cgff, _gcbcc)
				}
			}
		}
		for _afgc := 0; _afgc < len(_dggb); _afgc++ {
			_dbaf := _dggb[_afgc]
			for _, _aegd := range _dbaf.X().EG_RunInnerContent {
				if _aegd.FldChar != nil {
					switch _aegd.FldChar.FldCharTypeAttr {
					case wml.ST_FldCharTypeBegin:
						_dgac = _afgc
					case wml.ST_FldCharTypeSeparate:
						_edca = _afgc
					case wml.ST_FldCharTypeEnd:
						_bbde = _afgc
						if _dffb._gdfge != "" {
							_dffb._abdbd = _bcfe
							_dffb._bbcb = _dgac
							_dffb._gfaf = _bbde
							_dffb._cdcbe = _edca
							_cgff = append(_cgff, _dffb)
						}
						_dgac = -1
						_edca = -1
						_bbde = -1
						_dffb = mergeFieldInfo{}
					}
				} else if _aegd.InstrText != nil && strings.Contains(_aegd.InstrText.Content, "\u004d\u0045\u0052\u0047\u0045\u0046\u0049\u0045\u004c\u0044") {
					if _dgac != -1 && _bbde == -1 {
						_dffb = _cadg(_aegd.InstrText.Content)
					}
				}
			}
		}
	}
	return _cgff
}

// SetFirstColumn controls the conditional formatting for the first column in a table.
func (_ggac TableLook) SetFirstColumn(on bool) {
	if !on {
		_ggac.ctTblLook.FirstColumnAttr = &sharedTypes.ST_OnOff{}
		_ggac.ctTblLook.FirstColumnAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	} else {
		_ggac.ctTblLook.FirstColumnAttr = &sharedTypes.ST_OnOff{}
		_ggac.ctTblLook.FirstColumnAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	}
}
func (_efbf *Document) InsertTableAfter(relativeTo Paragraph) Table {
	return _efbf.insertTable(relativeTo, false)
}

// SetSize sets size attribute for a FormFieldTypeCheckBox in pt.
func (_caae FormField) SetSize(size uint64) {
	size *= 2
	if _caae._cbde.CheckBox != nil {
		_caae._cbde.CheckBox.Choice = wml.NewCT_FFCheckBoxChoice()
		_caae._cbde.CheckBox.Choice.Size = wml.NewCT_HpsMeasure()
		_caae._cbde.CheckBox.Choice.Size.ValAttr = wml.ST_HpsMeasure{ST_UnsignedDecimalNumber: &size}
	}
}

// SetStyle sets the table style name.
func (_bbgfc TableProperties) SetStyle(name string) {
	if name == "" {
		_bbgfc._efag.TblStyle = nil
	} else {
		_bbgfc._efag.TblStyle = wml.NewCT_String()
		_bbgfc._efag.TblStyle.ValAttr = name
	}
}

// SetStyle sets style to the text in watermark.
func (_ceed *WatermarkText) SetStyle(style vmldrawing.TextpathStyle) {
	_aedae := _ceed.getShape()
	if _ceed._bfbf != nil {
		_daagb := _ceed._bfbf.EG_ShapeElements
		if len(_daagb) > 0 && _daagb[0].Textpath != nil {
			var _bcba = style.String()
			_daagb[0].Textpath.StyleAttr = &_bcba
		}
		return
	}
	_gaad := _ceed.findNode(_aedae, "\u0074\u0065\u0078\u0074\u0070\u0061\u0074\u0068")
	for _dfafe, _dfgd := range _gaad.Attrs {
		if _dfgd.Name.Local == "\u0073\u0074\u0079l\u0065" {
			_gaad.Attrs[_dfafe].Value = style.String()
		}
	}
}

// SetThemeColor sets the color from the theme.
func (_gaae Color) SetThemeColor(t wml.ST_ThemeColor) { _gaae._ec.ThemeColorAttr = t }
func _dcbb(_agbd []*wml.EG_ContentBlockContent, _ggba *TableInfo) []TextItem {
	_fceg := []TextItem{}
	for _, _eaac := range _agbd {
		if _cdad := _eaac.Sdt; _cdad != nil {
			if _ccga := _cdad.SdtContent; _ccga != nil {
				_fceg = append(_fceg, _gdfd(_ccga.P, _ggba, nil)...)
			}
		}
		_fceg = append(_fceg, _gdfd(_eaac.P, _ggba, nil)...)
		for _, _cgda := range _eaac.Tbl {
			for _fegb, _aacb := range _cgda.EG_ContentRowContent {
				for _, _abacf := range _aacb.Tr {
					for _abaa, _agea := range _abacf.EG_ContentCellContent {
						for _, _bded := range _agea.Tc {
							_egbd := &TableInfo{Table: _cgda, Row: _abacf, Cell: _bded, RowIndex: _fegb, ColIndex: _abaa}
							for _, _ddaa := range _bded.EG_BlockLevelElts {
								_fceg = append(_fceg, _dcbb(_ddaa.EG_ContentBlockContent, _egbd)...)
							}
						}
					}
				}
			}
		}
	}
	return _fceg
}

// TextItem is used for keeping text with references to a paragraph and run or a table, a row and a cell where it is located.
type TextItem struct {
	Text        string
	DrawingInfo *DrawingInfo
	Paragraph   *wml.CT_P
	Hyperlink   *wml.CT_Hyperlink
	Run         *wml.CT_R
	TableInfo   *TableInfo
}

// Footnote returns the footnote based on the ID; this can be used nicely with
// the run.IsFootnote() functionality.
func (_gaf *Document) Footnote(id int64) Footnote {
	for _, _dgd := range _gaf.Footnotes() {
		if _dgd.id() == id {
			return _dgd
		}
	}
	return Footnote{}
}

// SetStyle sets the font size.
func (_dagcd RunProperties) SetStyle(style string) {
	if style == "" {
		_dagcd._gbdb.RStyle = nil
	} else {
		_dagcd._gbdb.RStyle = wml.NewCT_String()
		_dagcd._gbdb.RStyle.ValAttr = style
	}
}
func (_eada Endnote) id() int64 { return _eada._fagg.IdAttr }

// BoldValue returns the precise nature of the bold setting (unset, off or on).
func (_gaede RunProperties) BoldValue() OnOffValue { return _fgccb(_gaede._gbdb.B) }

// SetBehindDoc sets the behindDoc attribute of anchor.
func (_abc AnchoredDrawing) SetBehindDoc(val bool) { _abc._dgc.BehindDocAttr = val }

// InsertRowBefore inserts a row before another row
func (_ffeae Table) InsertRowBefore(r Row) Row {
	for _dcccc, _becg := range _ffeae.ctTbl.EG_ContentRowContent {
		if len(_becg.Tr) > 0 && r.X() == _becg.Tr[0] {
			_bfddg := wml.NewEG_ContentRowContent()
			_ffeae.ctTbl.EG_ContentRowContent = append(_ffeae.ctTbl.EG_ContentRowContent, nil)
			copy(_ffeae.ctTbl.EG_ContentRowContent[_dcccc+1:], _ffeae.ctTbl.EG_ContentRowContent[_dcccc:])
			_ffeae.ctTbl.EG_ContentRowContent[_dcccc] = _bfddg
			_aggea := wml.NewCT_Row()
			_bfddg.Tr = append(_bfddg.Tr, _aggea)
			return Row{_ffeae.doc, _aggea}
		}
	}
	return _ffeae.AddRow()
}

// X returns the inner wrapped XML type.
func (_df Bookmark) X() *wml.CT_Bookmark { return _df._gc }

// AddImage adds an image to the document package, returning a reference that
// can be used to add the image to a run and place it in the document contents.
func (_bbbd Header) AddImage(i common.Image) (common.ImageRef, error) {
	var _deeba common.Relationships
	for _eggde, _bbaa := range _bbbd._dbagd._geb {
		if _bbaa == _bbbd._deae {
			_deeba = _bbbd._dbagd._cbfd[_eggde]
		}
	}
	_bbga := common.MakeImageRef(i, &_bbbd._dbagd.DocBase, _deeba)
	if i.Data == nil && i.Path == "" {
		return _bbga, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0064\u0061t\u0061\u0020\u006f\u0072\u0020\u0061\u0020\u0070\u0061\u0074\u0068")
	}
	if i.Format == "" {
		return _bbga, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0061\u0020v\u0061\u006c\u0069\u0064\u0020\u0066\u006f\u0072\u006d\u0061\u0074")
	}
	if i.Size.X == 0 || i.Size.Y == 0 {
		return _bbga, errors.New("\u0069\u006d\u0061\u0067e\u0020\u006d\u0075\u0073\u0074\u0020\u0068\u0061\u0076\u0065 \u0061 \u0076\u0061\u006c\u0069\u0064\u0020\u0073i\u007a\u0065")
	}
	_bbbd._dbagd.Images = append(_bbbd._dbagd.Images, _bbga)
	_bafg := fmt.Sprintf("\u006d\u0065d\u0069\u0061\u002fi\u006d\u0061\u0067\u0065\u0025\u0064\u002e\u0025\u0073", len(_bbbd._dbagd.Images), i.Format)
	_dbbg := _deeba.AddRelationship(_bafg, unioffice.ImageType)
	_bbga.SetRelID(_dbbg.X().IdAttr)
	return _bbga, nil
}

// CellBorders are the borders for an individual
type CellBorders struct{ _gf *wml.CT_TcBorders }

// X returns the inner wrapped XML type.
func (_egee InlineDrawing) X() *wml.WdInline { return _egee._ecag }

// CharacterSpacingMeasure returns paragraph characters spacing with its measure which can be mm, cm, in, pt, pc or pi.
func (_aebecd ParagraphProperties) CharacterSpacingMeasure() string {
	if _bced := _aebecd._dfaf.RPr.Spacing; _bced != nil {
		_faeg := _bced.ValAttr
		if _faeg.ST_UniversalMeasure != nil {
			return *_faeg.ST_UniversalMeasure
		}
	}
	return ""
}
func (_cgecg *WatermarkPicture) getShapeType() *unioffice.XSDAny {
	return _cgecg.getInnerElement("\u0073h\u0061\u0070\u0065\u0074\u0079\u0070e")
}

// Value returns the tring value of a FormFieldTypeText or FormFieldTypeDropDown.
func (_bfbad FormField) Value() string {
	if _bfbad._cbde.TextInput != nil && _bfbad._gcbd.T != nil {
		return _bfbad._gcbd.T.Content
	} else if _bfbad._cbde.DdList != nil && _bfbad._cbde.DdList.Result != nil {
		_efaf := _bfbad.PossibleValues()
		_acbfd := int(_bfbad._cbde.DdList.Result.ValAttr)
		if _acbfd < len(_efaf) {
			return _efaf[_acbfd]
		}
	} else if _bfbad._cbde.CheckBox != nil {
		if _bfbad.IsChecked() {
			return "\u0074\u0072\u0075\u0065"
		}
		return "\u0066\u0061\u006cs\u0065"
	}
	return ""
}

// RunProperties controls run styling properties
type RunProperties struct{ _gbdb *wml.CT_RPr }

// ParagraphBorders allows manipulation of borders on a paragraph.
type ParagraphBorders struct {
	_babda *Document
	_fdge  *wml.CT_PBdr
}

// FormFieldType is the type of the form field.
//go:generate stringer -type=FormFieldType
type FormFieldType byte

// SetLayoutInCell sets the layoutInCell attribute of anchor.
func (_eb AnchoredDrawing) SetLayoutInCell(val bool) { _eb._dgc.LayoutInCellAttr = val }

// NewNumbering constructs a new numbering.
func NewNumbering() Numbering { _ffed := wml.NewNumbering(); return Numbering{_ffed} }

// SetWidth sets the cell width to a specified width.
func (_ggd CellProperties) SetWidth(d measurement.Distance) {
	_ggd._cgc.TcW = wml.NewCT_TblWidth()
	_ggd._cgc.TcW.TypeAttr = wml.ST_TblWidthDxa
	_ggd._cgc.TcW.WAttr = &wml.ST_MeasurementOrPercent{}
	_ggd._cgc.TcW.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_ggd._cgc.TcW.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(d / measurement.Twips))
}
func (_ggdce *WatermarkPicture) getInnerElement(_agdgc string) *unioffice.XSDAny {
	for _, _gaec := range _ggdce._cdff.Any {
		_dcbc, _bdefg := _gaec.(*unioffice.XSDAny)
		if _bdefg && (_dcbc.XMLName.Local == _agdgc || _dcbc.XMLName.Local == "\u0076\u003a"+_agdgc) {
			return _dcbc
		}
	}
	return nil
}

// SetThemeShade sets the shade based off the theme color.
func (_geda Color) SetThemeShade(s uint8) {
	_aab := fmt.Sprintf("\u0025\u0030\u0032\u0078", s)
	_geda._ec.ThemeShadeAttr = &_aab
}

// SetVAlignment sets the vertical alignment for an anchored image.
func (_fge AnchoredDrawing) SetVAlignment(v wml.WdST_AlignV) {
	_fge._dgc.PositionV.Choice = &wml.WdCT_PosVChoice{}
	_fge._dgc.PositionV.Choice.Align = v
}

// SetASCIITheme sets the font ASCII Theme.
func (_edfe Fonts) SetASCIITheme(t wml.ST_Theme) { _edfe._feae.AsciiThemeAttr = t }

// SetLineSpacing sets the spacing between lines in a paragraph.
func (_gafdf Paragraph) SetLineSpacing(d measurement.Distance, rule wml.ST_LineSpacingRule) {
	_gafdf.ensurePPr()
	if _gafdf._eagd.PPr.Spacing == nil {
		_gafdf._eagd.PPr.Spacing = wml.NewCT_Spacing()
	}
	_beeeb := _gafdf._eagd.PPr.Spacing
	if rule == wml.ST_LineSpacingRuleUnset {
		_beeeb.LineRuleAttr = wml.ST_LineSpacingRuleUnset
		_beeeb.LineAttr = nil
	} else {
		_beeeb.LineRuleAttr = rule
		_beeeb.LineAttr = &wml.ST_SignedTwipsMeasure{}
		_beeeb.LineAttr.Int64 = unioffice.Int64(int64(d / measurement.Twips))
	}
}

// GetStyleByID returns Style by it's IdAttr.
func (_ebdg *Document) GetStyleByID(id string) Style {
	for _, _cbda := range _ebdg.Styles._abca.Style {
		if _cbda.StyleIdAttr != nil && *_cbda.StyleIdAttr == id {
			return Style{_cbda}
		}
	}
	return Style{}
}

// Endnote is an individual endnote reference within the document.
type Endnote struct {
	_cceg *Document
	_fagg *wml.CT_FtnEdn
}

// Borders allows controlling individual cell borders.
func (_bbd CellProperties) Borders() CellBorders {
	if _bbd._cgc.TcBorders == nil {
		_bbd._cgc.TcBorders = wml.NewCT_TcBorders()
	}
	return CellBorders{_bbd._cgc.TcBorders}
}

// Cell is a table cell within a document (not a spreadsheet)
type Cell struct {
	_dga *Document
	_gge *wml.CT_Tc
}

// Emboss returns true if paragraph emboss is on.
func (_fbegd ParagraphProperties) Emboss() bool { return _cadf(_fbegd._dfaf.RPr.Emboss) }

// Index returns the index of the footer within the document.  This is used to
// form its zip packaged filename as well as to match it with its relationship
// ID.
func (_fgefc Footer) Index() int {
	for _faed, _egfe := range _fgefc._aegg._aba {
		if _egfe == _fgefc._fcc {
			return _faed
		}
	}
	return -1
}

// RemoveParagraph removes a paragraph from the endnote.
func (_fagb Endnote) RemoveParagraph(p Paragraph) {
	for _, _dced := range _fagb.content() {
		for _bfgc, _fdgg := range _dced.P {
			if _fdgg == p._eagd {
				copy(_dced.P[_bfgc:], _dced.P[_bfgc+1:])
				_dced.P = _dced.P[0 : len(_dced.P)-1]
				return
			}
		}
	}
}

// SetPossibleValues sets possible values for a FormFieldTypeDropDown.
func (_beff FormField) SetPossibleValues(values []string) {
	if _beff._cbde.DdList != nil {
		for _, _ggbb := range values {
			_gedac := wml.NewCT_String()
			_gedac.ValAttr = _ggbb
			_beff._cbde.DdList.ListEntry = append(_beff._cbde.DdList.ListEntry, _gedac)
		}
	}
}

// RemoveParagraph removes a paragraph from a footer.
func (_ccfa Footer) RemoveParagraph(p Paragraph) {
	for _, _bdgec := range _ccfa._fcc.EG_ContentBlockContent {
		for _fgbf, _cdfbf := range _bdgec.P {
			if _cdfbf == p._eagd {
				copy(_bdgec.P[_fgbf:], _bdgec.P[_fgbf+1:])
				_bdgec.P = _bdgec.P[0 : len(_bdgec.P)-1]
				return
			}
		}
	}
}

// SetKerning sets the run's font kerning.
func (_abfb RunProperties) SetKerning(size measurement.Distance) {
	_abfb._gbdb.Kern = wml.NewCT_HpsMeasure()
	_abfb._gbdb.Kern.ValAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(size / measurement.HalfPoint))
}

// SetLeft sets the left border to a specified type, color and thickness.
func (_adf CellBorders) SetLeft(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_adf._gf.Left = wml.NewCT_Border()
	_feadc(_adf._gf.Left, t, c, thickness)
}
func _bfafd(_eefg *wml.CT_P, _ebdef *wml.CT_Hyperlink, _bbdd *TableInfo, _eeab *DrawingInfo, _cagcf []*wml.EG_ContentRunContent) []TextItem {
	_dfcgb := []TextItem{}
	for _, _deec := range _cagcf {
		if _cadef := _deec.R; _cadef != nil {
			_ggec := bytes.NewBuffer([]byte{})
			for _, _bfba := range _cadef.EG_RunInnerContent {
				if _bfba.T != nil && _bfba.T.Content != "" {
					_ggec.WriteString(_bfba.T.Content)
				}
			}
			_dfcgb = append(_dfcgb, TextItem{Text: _ggec.String(), DrawingInfo: _eeab, Paragraph: _eefg, Hyperlink: _ebdef, Run: _cadef, TableInfo: _bbdd})
			for _, _ffbe := range _cadef.Extra {
				if _ffdc, _cgaeb := _ffbe.(*wml.AlternateContentRun); _cgaeb {
					_ddba := &DrawingInfo{Drawing: _ffdc.Choice.Drawing}
					for _, _ffgba := range _ddba.Drawing.Anchor {
						for _, _ffgd := range _ffgba.Graphic.GraphicData.Any {
							if _addfg, _ggdc := _ffgd.(*wml.WdWsp); _ggdc {
								if _addfg.WChoice != nil {
									if _gdec := _addfg.SpPr; _gdec != nil {
										if _gcea := _gdec.Xfrm; _gcea != nil {
											if _gfea := _gcea.Ext; _gfea != nil {
												_ddba.Width = _gfea.CxAttr
												_ddba.Height = _gfea.CyAttr
											}
										}
									}
									for _, _ecf := range _addfg.WChoice.Txbx.TxbxContent.EG_ContentBlockContent {
										_dfcgb = append(_dfcgb, _gdfd(_ecf.P, _bbdd, _ddba)...)
									}
								}
							}
						}
					}
				}
			}
		}
	}
	return _dfcgb
}

// TableConditionalFormatting controls the conditional formatting within a table
// style.
type TableConditionalFormatting struct{ _ecbge *wml.CT_TblStylePr }

// Endnotes returns the endnotes defined in the document.
func (_cffd *Document) Endnotes() []Endnote {
	_aebe := []Endnote{}
	for _, _gbcbb := range _cffd._ccb.CT_Endnotes.Endnote {
		_aebe = append(_aebe, Endnote{_cffd, _gbcbb})
	}
	return _aebe
}

// AddText adds tet to a run.
func (_cfcb Run) AddText(s string) {
	_fgga := wml.NewEG_RunInnerContent()
	_cfcb._adaad.EG_RunInnerContent = append(_cfcb._adaad.EG_RunInnerContent, _fgga)
	_fgga.T = wml.NewCT_Text()
	if unioffice.NeedsSpacePreserve(s) {
		_dacdg := "\u0070\u0072\u0065\u0073\u0065\u0072\u0076\u0065"
		_fgga.T.SpaceAttr = &_dacdg
	}
	_fgga.T.Content = s
}

// SetLeftIndent controls left indent of paragraph.
func (_bdae Paragraph) SetLeftIndent(m measurement.Distance) {
	_bdae.ensurePPr()
	_bfcfd := _bdae._eagd.PPr
	if _bfcfd.Ind == nil {
		_bfcfd.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_bfcfd.Ind.LeftAttr = nil
	} else {
		_bfcfd.Ind.LeftAttr = &wml.ST_SignedTwipsMeasure{}
		_bfcfd.Ind.LeftAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// TableLook returns the table look, or conditional formatting applied to a table style.
func (_ebba TableProperties) TableLook() TableLook {
	if _ebba._efag.TblLook == nil {
		_ebba._efag.TblLook = wml.NewCT_TblLook()
	}
	return TableLook{_ebba._efag.TblLook}
}

// CellProperties are a table cells properties within a document.
type CellProperties struct{ _cgc *wml.CT_TcPr }

// SizeMeasure returns font with its measure which can be mm, cm, in, pt, pc or pi.
func (_dadg RunProperties) SizeMeasure() string {
	if _fbea := _dadg._gbdb.Sz; _fbea != nil {
		_abeb := _fbea.ValAttr
		if _abeb.ST_PositiveUniversalMeasure != nil {
			return *_abeb.ST_PositiveUniversalMeasure
		}
	}
	return ""
}

// ComplexSizeValue returns the value of run font size for complex fonts in points.
func (_ggae RunProperties) ComplexSizeValue() float64 {
	if _bfaa := _ggae._gbdb.SzCs; _bfaa != nil {
		_fdfae := _bfaa.ValAttr
		if _fdfae.ST_UnsignedDecimalNumber != nil {
			return float64(*_fdfae.ST_UnsignedDecimalNumber) / 2
		}
	}
	return 0.0
}

// Headers returns the headers defined in the document.
func (_cab *Document) Headers() []Header {
	_ega := []Header{}
	for _, _gagf := range _cab._geb {
		_ega = append(_ega, Header{_cab, _gagf})
	}
	return _ega
}

// AddBookmark adds a bookmark to a document that can then be used from a hyperlink. Name is a document
// unique name that identifies the bookmark so it can be referenced from hyperlinks.
func (_gdef Paragraph) AddBookmark(name string) Bookmark {
	_acee := wml.NewEG_PContent()
	_acfgf := wml.NewEG_ContentRunContent()
	_acee.EG_ContentRunContent = append(_acee.EG_ContentRunContent, _acfgf)
	_ecgad := wml.NewEG_RunLevelElts()
	_acfgf.EG_RunLevelElts = append(_acfgf.EG_RunLevelElts, _ecgad)
	_bdcc := wml.NewEG_RangeMarkupElements()
	_daag := wml.NewCT_Bookmark()
	_bdcc.BookmarkStart = _daag
	_ecgad.EG_RangeMarkupElements = append(_ecgad.EG_RangeMarkupElements, _bdcc)
	_bdcc = wml.NewEG_RangeMarkupElements()
	_bdcc.BookmarkEnd = wml.NewCT_MarkupRange()
	_ecgad.EG_RangeMarkupElements = append(_ecgad.EG_RangeMarkupElements, _bdcc)
	_gdef._eagd.EG_PContent = append(_gdef._eagd.EG_PContent, _acee)
	_dfdg := Bookmark{_daag}
	_dfdg.SetName(name)
	return _dfdg
}

// InitializeDefault constructs the default styles.
func (_caecc Styles) InitializeDefault() {
	_caecc.initializeDocDefaults()
	_caecc.initializeStyleDefaults()
}

// InsertRunBefore inserts a run in the paragraph before the relative run.
func (_eegc Paragraph) InsertRunBefore(relativeTo Run) Run { return _eegc.insertRun(relativeTo, true) }

// GetSize return the size of anchor on the page.
func (_ffg AnchoredDrawing) GetSize() (_dc, _agd int64) {
	return _ffg._dgc.Extent.CxAttr, _ffg._dgc.Extent.CyAttr
}

// TableProperties are the properties for a table within a document
type TableProperties struct{ _efag *wml.CT_TblPr }

// Tables returns the tables defined in the document.
func (_fgd *Document) Tables() []Table {
	_baac := []Table{}
	if _fgd.doc.Body == nil {
		return nil
	}
	for _, _fffe := range _fgd.doc.Body.EG_BlockLevelElts {
		for _, _dae := range _fffe.EG_ContentBlockContent {
			for _, _fec := range _fgd.tables(_dae) {
				_baac = append(_baac, _fec)
			}
		}
	}
	return _baac
}

// X returns the inner wrapped XML type.
func (_eebgc Styles) X() *wml.Styles { return _eebgc._abca }

// ExtractTextOptions extraction text options.
type ExtractTextOptions struct {
	WithNumbering   bool
	NumberingIndent string
}

// SetAllCaps sets the run to all caps.
func (_dead RunProperties) SetAllCaps(b bool) {
	if !b {
		_dead._gbdb.Caps = nil
	} else {
		_dead._gbdb.Caps = wml.NewCT_OnOff()
	}
}

// RightToLeft returns true if paragraph text goes from right to left.
func (_fggf ParagraphProperties) RightToLeft() bool { return _cadf(_fggf._dfaf.RPr.Rtl) }

// GetTargetByRelId returns a target path with the associated relation ID in the
// document.
func (_dfde *Document) GetTargetByRelId(idAttr string) string {
	return _dfde._dab.GetTargetByRelId(idAttr)
}
func (_gbea Footnote) content() []*wml.EG_ContentBlockContent {
	var _ccbag []*wml.EG_ContentBlockContent
	for _, _bgga := range _gbea._bgcda.EG_BlockLevelElts {
		_ccbag = append(_ccbag, _bgga.EG_ContentBlockContent...)
	}
	return _ccbag
}

// X returns the inner wrapped XML type.
func (_acae Footer) X() *wml.Ftr { return _acae._fcc }

// SetEffect sets a text effect on the run.
func (_eeead RunProperties) SetEffect(e wml.ST_TextEffect) {
	if e == wml.ST_TextEffectUnset {
		_eeead._gbdb.Effect = nil
	} else {
		_eeead._gbdb.Effect = wml.NewCT_TextEffect()
		_eeead._gbdb.Effect.ValAttr = wml.ST_TextEffectShimmer
	}
}

// SetTargetBookmark sets the bookmark target of the hyperlink.
func (_ccfe HyperLink) SetTargetBookmark(bm Bookmark) {
	_ccfe._baaf.AnchorAttr = unioffice.String(bm.Name())
	_ccfe._baaf.IdAttr = nil
}

// X returns the inner wrapped XML type.
func (_cede HyperLink) X() *wml.CT_Hyperlink { return _cede._baaf }

// FindNodeByStyleName return slice of node base on style name.
func (_cbbfg *Nodes) FindNodeByStyleName(styleName string) []Node {
	_eedg := []Node{}
	for _, _dfggb := range _cbbfg._gabfc {
		switch _becf := _dfggb._ggda.(type) {
		case *Paragraph:
			if _becf != nil {
				if _dbgad, _cfde := _dfggb._cdbd.Styles.SearchStyleByName(styleName); _cfde {
					_bbbe := _becf.Style()
					if _bbbe == _dbgad.StyleID() {
						_eedg = append(_eedg, _dfggb)
					}
				}
			}
		case *Table:
			if _becf != nil {
				if _aecc, _dfea := _dfggb._cdbd.Styles.SearchStyleByName(styleName); _dfea {
					_cdcf := _becf.Style()
					if _cdcf == _aecc.StyleID() {
						_eedg = append(_eedg, _dfggb)
					}
				}
			}
		}
		_afde := Nodes{_gabfc: _dfggb.Children}
		_eedg = append(_eedg, _afde.FindNodeByStyleName(styleName)...)
	}
	return _eedg
}

// SetLineSpacing sets the spacing between lines in a paragraph.
func (_gcdb ParagraphSpacing) SetLineSpacing(d measurement.Distance, rule wml.ST_LineSpacingRule) {
	if rule == wml.ST_LineSpacingRuleUnset {
		_gcdb._ffede.LineRuleAttr = wml.ST_LineSpacingRuleUnset
		_gcdb._ffede.LineAttr = nil
	} else {
		_gcdb._ffede.LineRuleAttr = rule
		_gcdb._ffede.LineAttr = &wml.ST_SignedTwipsMeasure{}
		_gcdb._ffede.LineAttr.Int64 = unioffice.Int64(int64(d / measurement.Twips))
	}
}

// SetName sets the name of the image, visible in the properties of the image
// within Word.
func (_bc AnchoredDrawing) SetName(name string) {
	_bc._dgc.DocPr.NameAttr = name
	for _, _ff := range _bc._dgc.Graphic.GraphicData.Any {
		if _ab, _db := _ff.(*picture.Pic); _db {
			_ab.NvPicPr.CNvPr.DescrAttr = unioffice.String(name)
		}
	}
}

// DrawingInfo is used for keep information about a drawing wrapping a textbox where the text is located.
type DrawingInfo struct {
	Drawing *wml.CT_Drawing
	Width   int64
	Height  int64
}

// AddCell adds a cell to a row and returns it
func (row Row) AddCell() Cell {
	cellContent := wml.NewEG_ContentCellContent()
	row.ctRow.EG_ContentCellContent = append(row.ctRow.EG_ContentCellContent, cellContent)
	tc := wml.NewCT_Tc()
	cellContent.Tc = append(cellContent.Tc, tc)
	return Cell{row.doc, tc}
}

// SetHorizontalBanding controls the conditional formatting for horizontal banding.
func (tl TableLook) SetHorizontalBanding(on bool) {
	if !on {
		tl.ctTblLook.NoHBandAttr = &sharedTypes.ST_OnOff{}
		tl.ctTblLook.NoHBandAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	} else {
		tl.ctTblLook.NoHBandAttr = &sharedTypes.ST_OnOff{}
		tl.ctTblLook.NoHBandAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	}
}

// InlineDrawing is an inlined image within a run.
type InlineDrawing struct {
	_aaaa *Document
	_ecag *wml.WdInline
}

func _adaa(_ddg *wml.CT_Tbl, _gab *wml.CT_P, _bbef bool) *wml.CT_Tbl {
	for _, _faa := range _ddg.EG_ContentRowContent {
		for _, _dgcd := range _faa.Tr {
			for _, _baec := range _dgcd.EG_ContentCellContent {
				for _, _egd := range _baec.Tc {
					for _efgf, _ffc := range _egd.EG_BlockLevelElts {
						for _, _gda := range _ffc.EG_ContentBlockContent {
							for _gdd, _gfgf := range _gda.P {
								if _gfgf == _gab {
									_bgfd := wml.NewEG_BlockLevelElts()
									_acg := wml.NewEG_ContentBlockContent()
									_bgfd.EG_ContentBlockContent = append(_bgfd.EG_ContentBlockContent, _acg)
									_bagd := wml.NewCT_Tbl()
									_acg.Tbl = append(_acg.Tbl, _bagd)
									_egd.EG_BlockLevelElts = append(_egd.EG_BlockLevelElts, nil)
									if _bbef {
										copy(_egd.EG_BlockLevelElts[_efgf+1:], _egd.EG_BlockLevelElts[_efgf:])
										_egd.EG_BlockLevelElts[_efgf] = _bgfd
										if _gdd != 0 {
											_dac := wml.NewEG_BlockLevelElts()
											_ggde := wml.NewEG_ContentBlockContent()
											_dac.EG_ContentBlockContent = append(_dac.EG_ContentBlockContent, _ggde)
											_ggde.P = _gda.P[:_gdd]
											_egd.EG_BlockLevelElts = append(_egd.EG_BlockLevelElts, nil)
											copy(_egd.EG_BlockLevelElts[_efgf+1:], _egd.EG_BlockLevelElts[_efgf:])
											_egd.EG_BlockLevelElts[_efgf] = _dac
										}
										_gda.P = _gda.P[_gdd:]
									} else {
										copy(_egd.EG_BlockLevelElts[_efgf+2:], _egd.EG_BlockLevelElts[_efgf+1:])
										_egd.EG_BlockLevelElts[_efgf+1] = _bgfd
										if _gdd != len(_gda.P)-1 {
											_cdfe := wml.NewEG_BlockLevelElts()
											_caa := wml.NewEG_ContentBlockContent()
											_cdfe.EG_ContentBlockContent = append(_cdfe.EG_ContentBlockContent, _caa)
											_caa.P = _gda.P[_gdd+1:]
											_egd.EG_BlockLevelElts = append(_egd.EG_BlockLevelElts, nil)
											copy(_egd.EG_BlockLevelElts[_efgf+3:], _egd.EG_BlockLevelElts[_efgf+2:])
											_egd.EG_BlockLevelElts[_efgf+2] = _cdfe
										} else {
											_aac := wml.NewEG_BlockLevelElts()
											_dfg := wml.NewEG_ContentBlockContent()
											_aac.EG_ContentBlockContent = append(_aac.EG_ContentBlockContent, _dfg)
											_dfg.P = []*wml.CT_P{wml.NewCT_P()}
											_egd.EG_BlockLevelElts = append(_egd.EG_BlockLevelElts, nil)
											copy(_egd.EG_BlockLevelElts[_efgf+3:], _egd.EG_BlockLevelElts[_efgf+2:])
											_egd.EG_BlockLevelElts[_efgf+2] = _aac
										}
										_gda.P = _gda.P[:_gdd+1]
									}
									return _bagd
								}
							}
							for _, _bcd := range _gda.Tbl {
								_ddad := _adaa(_bcd, _gab, _bbef)
								if _ddad != nil {
									return _ddad
								}
							}
						}
					}
				}
			}
		}
	}
	return nil
}

// SetCellSpacing sets the cell spacing within a table.
func (_efade TableProperties) SetCellSpacing(m measurement.Distance) {
	_efade._efag.TblCellSpacing = wml.NewCT_TblWidth()
	_efade._efag.TblCellSpacing.TypeAttr = wml.ST_TblWidthDxa
	_efade._efag.TblCellSpacing.WAttr = &wml.ST_MeasurementOrPercent{}
	_efade._efag.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_efade._efag.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(m / measurement.Dxa))
}

// Properties returns the run properties.
func (_cfff Run) Properties() RunProperties {
	if _cfff._adaad.RPr == nil {
		_cfff._adaad.RPr = wml.NewCT_RPr()
	}
	return RunProperties{_cfff._adaad.RPr}
}

// RemoveParagraph removes a paragraph from a document.
func (_efgd *Document) RemoveParagraph(p Paragraph) {
	if _efgd.doc.Body == nil {
		return
	}
	for _, _agg := range _efgd.doc.Body.EG_BlockLevelElts {
		for _, _cddf := range _agg.EG_ContentBlockContent {
			for _deb, _cgde := range _cddf.P {
				if _cgde == p._eagd {
					copy(_cddf.P[_deb:], _cddf.P[_deb+1:])
					_cddf.P = _cddf.P[0 : len(_cddf.P)-1]
					return
				}
			}
			if _cddf.Sdt != nil && _cddf.Sdt.SdtContent != nil && _cddf.Sdt.SdtContent.P != nil {
				for _bgaf, _bfgb := range _cddf.Sdt.SdtContent.P {
					if _bfgb == p._eagd {
						copy(_cddf.P[_bgaf:], _cddf.P[_bgaf+1:])
						_cddf.P = _cddf.P[0 : len(_cddf.P)-1]
						return
					}
				}
			}
		}
	}
	for _, _eadd := range _efgd.Tables() {
		for _, _dge := range _eadd.Rows() {
			for _, _gdfb := range _dge.Cells() {
				for _, _cfb := range _gdfb._gge.EG_BlockLevelElts {
					for _, _fdag := range _cfb.EG_ContentBlockContent {
						for _bce, _ggab := range _fdag.P {
							if _ggab == p._eagd {
								copy(_fdag.P[_bce:], _fdag.P[_bce+1:])
								_fdag.P = _fdag.P[0 : len(_fdag.P)-1]
								return
							}
						}
					}
				}
			}
		}
	}
	for _, _abcg := range _efgd.Headers() {
		_abcg.RemoveParagraph(p)
	}
	for _, _gee := range _efgd.Footers() {
		_gee.RemoveParagraph(p)
	}
}

// Paragraphs returns the paragraphs within a structured document tag.
func (_ecbb StructuredDocumentTag) Paragraphs() []Paragraph {
	if _ecbb._afadb.SdtContent == nil {
		return nil
	}
	_cabf := []Paragraph{}
	for _, _gcfee := range _ecbb._afadb.SdtContent.P {
		_cabf = append(_cabf, Paragraph{_ecbb._fdad, _gcfee})
	}
	return _cabf
}

// ClearContent clears any content in the run (text, tabs, breaks, etc.)
func (_aadd Run) ClearContent() { _aadd._adaad.EG_RunInnerContent = nil }

// X returns the inner wrapped XML type.
func (_gfebg NumberingDefinition) X() *wml.CT_AbstractNum { return _gfebg._agff }
func _afa() (*dml.CT_Point2D, []*dml.CT_Point2D) {
	var (
		_fa  int64 = 0
		_aag int64 = 21600
	)
	_fe := dml.ST_Coordinate{ST_CoordinateUnqualified: &_fa, ST_UniversalMeasure: nil}
	_agc := dml.ST_Coordinate{ST_CoordinateUnqualified: &_aag, ST_UniversalMeasure: nil}
	_afd := dml.NewCT_Point2D()
	_afd.XAttr = _fe
	_afd.YAttr = _fe
	_gef := []*dml.CT_Point2D{&dml.CT_Point2D{XAttr: _fe, YAttr: _agc}, &dml.CT_Point2D{XAttr: _agc, YAttr: _agc}, &dml.CT_Point2D{XAttr: _agc, YAttr: _fe}, _afd}
	return _afd, _gef
}

// SetSpacing sets the spacing that comes before and after the paragraph.
func (_ecbgd ParagraphStyleProperties) SetSpacing(before, after measurement.Distance) {
	if _ecbgd._gfee.Spacing == nil {
		_ecbgd._gfee.Spacing = wml.NewCT_Spacing()
	}
	if before == measurement.Zero {
		_ecbgd._gfee.Spacing.BeforeAttr = nil
	} else {
		_ecbgd._gfee.Spacing.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
		_ecbgd._gfee.Spacing.BeforeAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(before / measurement.Twips))
	}
	if after == measurement.Zero {
		_ecbgd._gfee.Spacing.AfterAttr = nil
	} else {
		_ecbgd._gfee.Spacing.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
		_ecbgd._gfee.Spacing.AfterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(after / measurement.Twips))
	}
}
func (_gbcb *chart) Target() string { return _gbcb._cce }

// Endnote returns the endnote based on the ID; this can be used nicely with
// the run.IsEndnote() functionality.
func (_eedb *Document) Endnote(id int64) Endnote {
	for _, _cbcg := range _eedb.Endnotes() {
		if _cbcg.id() == id {
			return _cbcg
		}
	}
	return Endnote{}
}
func _cdbee() *vml.Textpath {
	_aaaae := vml.NewTextpath()
	_gfcbg := "\u0066\u006f\u006e\u0074\u002d\u0066\u0061\u006d\u0069l\u0079\u003a\u0022\u0043\u0061\u006c\u0069\u0062\u0072\u0069\u0022\u003b\u0066\u006f\u006e\u0074\u002d\u0073\u0069\u007a\u0065\u003a\u00366\u0070\u0074;\u0066\u006fn\u0074\u002d\u0077\u0065\u0069\u0067\u0068\u0074\u003a\u0062\u006f\u006c\u0064;f\u006f\u006e\u0074\u002d\u0073\u0074\u0079\u006c\u0065:\u0069\u0074\u0061\u006c\u0069\u0063"
	_aaaae.StyleAttr = &_gfcbg
	_bbca := "\u0041\u0053\u0041\u0050"
	_aaaae.StringAttr = &_bbca
	return _aaaae
}

// AddHyperLink adds a new hyperlink to a parapgraph.
func (_gageb Paragraph) AddHyperLink() HyperLink {
	_eeebe := wml.NewEG_PContent()
	_gageb._eagd.EG_PContent = append(_gageb._eagd.EG_PContent, _eeebe)
	_eeebe.Hyperlink = wml.NewCT_Hyperlink()
	return HyperLink{_gageb._fagf, _eeebe.Hyperlink}
}

// Index returns the index of the header within the document.  This is used to
// form its zip packaged filename as well as to match it with its relationship
// ID.
func (_bgba Header) Index() int {
	for _dbb, _gafc := range _bgba._dbagd._geb {
		if _gafc == _bgba._deae {
			return _dbb
		}
	}
	return -1
}

// NewStyles constructs a new empty Styles
func NewStyles() Styles { return Styles{wml.NewStyles()} }
func (_cdbg *Document) insertImageFromNode(_ceec Node) {
	for _, _bdba := range _ceec.AnchoredDrawings {
		if _gefb, _eegga := _bdba.GetImage(); _eegga {
			_geae, _gaab := common.ImageFromFile(_gefb.Path())
			if _gaab != nil {
				logger.Log.Debug("\u0075\u006e\u0061\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0063r\u0065\u0061\u0074\u0065\u0020\u0069\u006d\u0061\u0067\u0065:\u0020\u0025\u0073", _gaab)
			}
			_dgaef, _gaab := _cdbg.AddImage(_geae)
			if _gaab != nil {
				logger.Log.Debug("u\u006e\u0061\u0062\u006c\u0065\u0020t\u006f\u0020\u0061\u0064\u0064\u0020i\u006d\u0061\u0067\u0065\u0020\u0074\u006f \u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u003a\u0020%\u0073", _gaab)
			}
			_gdcf := _cdbg._dab.GetByRelId(_dgaef.RelID())
			_gdcf.SetID(_gefb.RelID())
		}
	}
	for _, _agaf := range _ceec.InlineDrawings {
		if _gded, _baab := _agaf.GetImage(); _baab {
			_ffee, _gdeg := common.ImageFromFile(_gded.Path())
			if _gdeg != nil {
				logger.Log.Debug("\u0075\u006e\u0061\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0063r\u0065\u0061\u0074\u0065\u0020\u0069\u006d\u0061\u0067\u0065:\u0020\u0025\u0073", _gdeg)
			}
			_eafc, _gdeg := _cdbg.AddImage(_ffee)
			if _gdeg != nil {
				logger.Log.Debug("u\u006e\u0061\u0062\u006c\u0065\u0020t\u006f\u0020\u0061\u0064\u0064\u0020i\u006d\u0061\u0067\u0065\u0020\u0074\u006f \u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u003a\u0020%\u0073", _gdeg)
			}
			_addf := _cdbg._dab.GetByRelId(_eafc.RelID())
			_addf.SetID(_gded.RelID())
		}
	}
}

// Name returns the name of the field.
func (_dfge FormField) Name() string { return *_dfge._cbde.Name[0].ValAttr }

// SetUpdateFieldsOnOpen controls if fields are recalculated upon opening the
// document. This is useful for things like a table of contents as the library
// only adds the field code and relies on Word/LibreOffice to actually compute
// the content.
func (_gcbg Settings) SetUpdateFieldsOnOpen(b bool) {
	if !b {
		_gcbg._cdbbf.UpdateFields = nil
	} else {
		_gcbg._cdbbf.UpdateFields = wml.NewCT_OnOff()
	}
}
func (_eeea Paragraph) addFldCharsForField(_ceaff, _agag string) FormField {
	_feaccb := _eeea.addBeginFldChar(_ceaff)
	_ffdfb := FormField{_cbde: _feaccb}
	_cbge := _eeea._fagf.Bookmarks()
	_gbcg := int64(len(_cbge))
	if _ceaff != "" {
		_eeea.addStartBookmark(_gbcg, _ceaff)
	}
	_eeea.addInstrText(_agag)
	_eeea.addSeparateFldChar()
	if _agag == "\u0046\u004f\u0052\u004d\u0054\u0045\u0058\u0054" {
		_ecfc := _eeea.AddRun()
		_ebbd := wml.NewEG_RunInnerContent()
		_ecfc._adaad.EG_RunInnerContent = []*wml.EG_RunInnerContent{_ebbd}
		_ffdfb._gcbd = _ebbd
	}
	_eeea.addEndFldChar()
	if _ceaff != "" {
		_eeea.addEndBookmark(_gbcg)
	}
	return _ffdfb
}

// X returns the inner wrapped XML type.
func (_cedd TableWidth) X() *wml.CT_TblWidth { return _cedd._egbb }

// SetCellSpacingAuto sets the cell spacing within a table to automatic.
func (_fbbad TableStyleProperties) SetCellSpacingAuto() {
	_fbbad._degc.TblCellSpacing = wml.NewCT_TblWidth()
	_fbbad._degc.TblCellSpacing.TypeAttr = wml.ST_TblWidthAuto
}

// SetPageBreakBefore controls if there is a page break before this paragraph.
func (_dcbaa ParagraphProperties) SetPageBreakBefore(b bool) {
	if !b {
		_dcbaa._dfaf.PageBreakBefore = nil
	} else {
		_dcbaa._dfaf.PageBreakBefore = wml.NewCT_OnOff()
	}
}

// Bookmark is a bookmarked location within a document that can be referenced
// with a hyperlink.
type Bookmark struct{ _gc *wml.CT_Bookmark }

// ParagraphProperties returns the paragraph properties controlling text formatting within the table.
func (_feda TableConditionalFormatting) ParagraphProperties() ParagraphStyleProperties {
	if _feda._ecbge.PPr == nil {
		_feda._ecbge.PPr = wml.NewCT_PPrGeneral()
	}
	return ParagraphStyleProperties{_feda._ecbge.PPr}
}

// SetDoubleStrikeThrough sets the run to double strike-through.
func (_ecfb RunProperties) SetDoubleStrikeThrough(b bool) {
	if !b {
		_ecfb._gbdb.Dstrike = nil
	} else {
		_ecfb._gbdb.Dstrike = wml.NewCT_OnOff()
	}
}

// Paragraphs returns the paragraphs defined in a footnote.
func (_fbca Footnote) Paragraphs() []Paragraph {
	_agbb := []Paragraph{}
	for _, _bfbea := range _fbca.content() {
		for _, _aaf := range _bfbea.P {
			_agbb = append(_agbb, Paragraph{_fbca._gffg, _aaf})
		}
	}
	return _agbb
}

// Validate validates the structure and in cases where it't possible, the ranges
// of elements within a document. A validation error dones't mean that the
// document won't work in MS Word or LibreOffice, but it's worth checking into.
func (_geef *Document) Validate() error {
	if _geef == nil || _geef.doc == nil {
		return errors.New("\u0064o\u0063\u0075m\u0065\u006e\u0074\u0020n\u006f\u0074\u0020i\u006e\u0069\u0074\u0069\u0061\u006c\u0069\u007a\u0065d \u0063\u006f\u0072r\u0065\u0063t\u006c\u0079\u002c\u0020\u006e\u0069l\u0020\u0062a\u0073\u0065")
	}
	for _, _bbge := range []func() error{_geef.validateTableCells, _geef.validateBookmarks} {
		if _bffb := _bbge(); _bffb != nil {
			return _bffb
		}
	}
	if _adef := _geef.doc.Validate(); _adef != nil {
		return _adef
	}
	return nil
}
func _aebg(_dfag *wml.EG_ContentBlockContent) []Bookmark {
	_cfea := []Bookmark{}
	for _, _ggag := range _dfag.P {
		for _, _fcgc := range _ggag.EG_PContent {
			for _, _dcfc := range _fcgc.EG_ContentRunContent {
				for _, _cdbb := range _dcfc.EG_RunLevelElts {
					for _, _ggee := range _cdbb.EG_RangeMarkupElements {
						if _ggee.BookmarkStart != nil {
							_cfea = append(_cfea, Bookmark{_ggee.BookmarkStart})
						}
					}
				}
			}
		}
	}
	for _, _dea := range _dfag.EG_RunLevelElts {
		for _, _gefe := range _dea.EG_RangeMarkupElements {
			if _gefe.BookmarkStart != nil {
				_cfea = append(_cfea, Bookmark{_gefe.BookmarkStart})
			}
		}
	}
	for _, _cfbdb := range _dfag.Tbl {
		for _, _dcba := range _cfbdb.EG_ContentRowContent {
			for _, _bbag := range _dcba.Tr {
				for _, _decfd := range _bbag.EG_ContentCellContent {
					for _, _agdg := range _decfd.Tc {
						for _, _daac := range _agdg.EG_BlockLevelElts {
							for _, _dee := range _daac.EG_ContentBlockContent {
								for _, _dff := range _aebg(_dee) {
									_cfea = append(_cfea, _dff)
								}
							}
						}
					}
				}
			}
		}
	}
	return _cfea
}

// SetPageMargins sets the page margins for a section
func (_efcc Section) SetPageMargins(top, right, bottom, left, header, footer, gutter measurement.Distance) {
	_dfeag := wml.NewCT_PageMar()
	_dfeag.TopAttr.Int64 = unioffice.Int64(int64(top / measurement.Twips))
	_dfeag.BottomAttr.Int64 = unioffice.Int64(int64(bottom / measurement.Twips))
	_dfeag.RightAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(right / measurement.Twips))
	_dfeag.LeftAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(left / measurement.Twips))
	_dfeag.HeaderAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(header / measurement.Twips))
	_dfeag.FooterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(footer / measurement.Twips))
	_dfeag.GutterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(gutter / measurement.Twips))
	_efcc._ddcag.PgMar = _dfeag
}

// Spacing returns the paragraph spacing settings.
func (_facb ParagraphProperties) Spacing() ParagraphSpacing {
	if _facb._dfaf.Spacing == nil {
		_facb._dfaf.Spacing = wml.NewCT_Spacing()
	}
	return ParagraphSpacing{_facb._dfaf.Spacing}
}

// Node is document element node,
// contains Paragraph or Table element.
type Node struct {
	_cdbd            *Document
	_ggda            interface{}
	Style            Style
	AnchoredDrawings []AnchoredDrawing
	InlineDrawings   []InlineDrawing
	Children         []Node
}

// SetSize sets the size of the displayed image on the page.
func (_bbe AnchoredDrawing) SetSize(w, h measurement.Distance) {
	_bbe._dgc.Extent.CxAttr = int64(float64(w*measurement.Pixel72) / measurement.EMU)
	_bbe._dgc.Extent.CyAttr = int64(float64(h*measurement.Pixel72) / measurement.EMU)
}

// SetLeft sets the left border to a specified type, color and thickness.
func (_faace ParagraphBorders) SetLeft(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_faace._fdge.Left = wml.NewCT_Border()
	_bbgf(_faace._fdge.Left, t, c, thickness)
}

// SetLayout controls the table layout. wml.ST_TblLayoutTypeAutofit corresponds
// to "Automatically resize to fit contents" being checked, while
// wml.ST_TblLayoutTypeFixed corresponds to it being unchecked.
func (_cgbd TableProperties) SetLayout(l wml.ST_TblLayoutType) {
	if l == wml.ST_TblLayoutTypeUnset || l == wml.ST_TblLayoutTypeAutofit {
		_cgbd._efag.TblLayout = nil
	} else {
		_cgbd._efag.TblLayout = wml.NewCT_TblLayoutType()
		_cgbd._efag.TblLayout.TypeAttr = l
	}
}

type chart struct {
	_ffb *dmlChart.ChartSpace
	_fda string
	_cce string
}

// GetImage returns the ImageRef associated with an InlineDrawing.
func (_adace InlineDrawing) GetImage() (common.ImageRef, bool) {
	_cgdf := _adace._ecag.Graphic.GraphicData.Any
	if len(_cgdf) > 0 {
		_debg, _ffaf := _cgdf[0].(*picture.Pic)
		if _ffaf {
			if _debg.BlipFill != nil && _debg.BlipFill.Blip != nil && _debg.BlipFill.Blip.EmbedAttr != nil {
				return _adace._aaaa.GetImageByRelID(*_debg.BlipFill.Blip.EmbedAttr)
			}
		}
	}
	return common.ImageRef{}, false
}

// SetNumberingDefinition sets the numbering definition ID via a NumberingDefinition
// defined in numbering.xml
func (_feff Paragraph) SetNumberingDefinition(nd NumberingDefinition) {
	_feff.ensurePPr()
	if _feff._eagd.PPr.NumPr == nil {
		_feff._eagd.PPr.NumPr = wml.NewCT_NumPr()
	}
	_aaafg := wml.NewCT_DecimalNumber()
	_aebec := int64(-1)
	for _, _acbef := range _feff._fagf.Numbering._cbag.Num {
		if _acbef.AbstractNumId != nil && _acbef.AbstractNumId.ValAttr == nd.AbstractNumberID() {
			_aebec = _acbef.NumIdAttr
		}
	}
	if _aebec == -1 {
		_ffec := wml.NewCT_Num()
		_feff._fagf.Numbering._cbag.Num = append(_feff._fagf.Numbering._cbag.Num, _ffec)
		_ffec.NumIdAttr = int64(len(_feff._fagf.Numbering._cbag.Num))
		_ffec.AbstractNumId = wml.NewCT_DecimalNumber()
		_ffec.AbstractNumId.ValAttr = nd.AbstractNumberID()
	}
	_aaafg.ValAttr = _aebec
	_feff._eagd.PPr.NumPr.NumId = _aaafg
}

// GetText returns text in the watermark.
func (_gcaad *WatermarkText) GetText() string {
	_geedc := _gcaad.getShape()
	if _gcaad._bfbf != nil {
		_ffbee := _gcaad._bfbf.EG_ShapeElements
		if len(_ffbee) > 0 && _ffbee[0].Textpath != nil {
			return *_ffbee[0].Textpath.StringAttr
		}
	} else {
		_ccab := _gcaad.findNode(_geedc, "\u0074\u0065\u0078\u0074\u0070\u0061\u0074\u0068")
		for _, _aced := range _ccab.Attrs {
			if _aced.Name.Local == "\u0073\u0074\u0072\u0069\u006e\u0067" {
				return _aced.Value
			}
		}
	}
	return ""
}
func (_cbcb Paragraph) addInstrText(_dcfb string) *wml.CT_Text {
	_gffd := _cbcb.AddRun()
	_fgdg := _gffd.X()
	_bgfef := wml.NewEG_RunInnerContent()
	_gbfa := wml.NewCT_Text()
	_egedb := "\u0070\u0072\u0065\u0073\u0065\u0072\u0076\u0065"
	_gbfa.SpaceAttr = &_egedb
	_gbfa.Content = "\u0020" + _dcfb + "\u0020"
	_bgfef.InstrText = _gbfa
	_fgdg.EG_RunInnerContent = append(_fgdg.EG_RunInnerContent, _bgfef)
	return _gbfa
}

// MultiLevelType returns the multilevel type, or ST_MultiLevelTypeUnset if not set.
func (_dffed NumberingDefinition) MultiLevelType() wml.ST_MultiLevelType {
	if _dffed._agff.MultiLevelType != nil {
		return _dffed._agff.MultiLevelType.ValAttr
	} else {
		return wml.ST_MultiLevelTypeUnset
	}
}

// DoubleStrike returns true if run is double striked.
func (_fffeg RunProperties) DoubleStrike() bool { return _cadf(_fffeg._gbdb.Dstrike) }

// IsBold returns true if the run has been set to bold.
func (_gcaba RunProperties) IsBold() bool { return _gcaba.BoldValue() == OnOffValueOn }

// FindNodeByStyleId return slice of node base on style id.
func (_gddbb *Nodes) FindNodeByStyleId(styleId string) []Node {
	_edbdd := []Node{}
	for _, _ccfff := range _gddbb._gabfc {
		switch _ccda := _ccfff._ggda.(type) {
		case *Paragraph:
			if _ccda != nil && _ccda.Style() == styleId {
				_edbdd = append(_edbdd, _ccfff)
			}
		case *Table:
			if _ccda != nil && _ccda.Style() == styleId {
				_edbdd = append(_edbdd, _ccfff)
			}
		}
		_fcbb := Nodes{_gabfc: _ccfff.Children}
		_edbdd = append(_edbdd, _fcbb.FindNodeByStyleId(styleId)...)
	}
	return _edbdd
}

// SearchStylesById returns style by its id.
func (_gacbae Styles) SearchStyleById(id string) (Style, bool) {
	for _, _bbgg := range _gacbae._abca.Style {
		if _bbgg.StyleIdAttr != nil {
			if *_bbgg.StyleIdAttr == id {
				return Style{_bbgg}, true
			}
		}
	}
	return Style{}, false
}

// SetHANSITheme sets the font H ANSI Theme.
func (_gggb Fonts) SetHANSITheme(t wml.ST_Theme) { _gggb._feae.HAnsiThemeAttr = t }

// SetLeft sets the cell left margin
func (_cff CellMargins) SetLeft(d measurement.Distance) {
	_cff._cdae.Left = wml.NewCT_TblWidth()
	_age(_cff._cdae.Left, d)
}

// SetAlignment controls the paragraph alignment
func (_fdbgc ParagraphStyleProperties) SetAlignment(align wml.ST_Jc) {
	if align == wml.ST_JcUnset {
		_fdbgc._gfee.Jc = nil
	} else {
		_fdbgc._gfee.Jc = wml.NewCT_Jc()
		_fdbgc._gfee.Jc.ValAttr = align
	}
}

// SetFontFamily sets the Ascii & HAnsi fonly family for a run.
func (_fcgcd RunProperties) SetFontFamily(family string) {
	if _fcgcd._gbdb.RFonts == nil {
		_fcgcd._gbdb.RFonts = wml.NewCT_Fonts()
	}
	_fcgcd._gbdb.RFonts.AsciiAttr = unioffice.String(family)
	_fcgcd._gbdb.RFonts.HAnsiAttr = unioffice.String(family)
	_fcgcd._gbdb.RFonts.EastAsiaAttr = unioffice.String(family)
}

// RunProperties returns the RunProperties controlling numbering level font, etc.
func (_afed NumberingLevel) RunProperties() RunProperties {
	if _afed.lvl.RPr == nil {
		_afed.lvl.RPr = wml.NewCT_RPr()
	}
	return RunProperties{_afed.lvl.RPr}
}
func _bdcf() *vml.Handles {
	_bfbg := vml.NewHandles()
	_acdg := vml.NewCT_H()
	_afbd := "\u0023\u0030\u002c\u0062\u006f\u0074\u0074\u006f\u006dR\u0069\u0067\u0068\u0074"
	_acdg.PositionAttr = &_afbd
	_dfffd := "\u0036\u0036\u0032\u0039\u002c\u0031\u0034\u0039\u0037\u0031"
	_acdg.XrangeAttr = &_dfffd
	_bfbg.H = []*vml.CT_H{_acdg}
	return _bfbg
}

// HasEndnotes returns a bool based on the presence or abscence of endnotes within
// the document.
func (_dde *Document) HasEndnotes() bool { return _dde._ccb != nil }

// AddDrawingAnchored adds an anchored (floating) drawing from an ImageRef.
func (_ffacf Run) AddDrawingAnchored(img common.ImageRef) (AnchoredDrawing, error) {
	_dcbae := _ffacf.newIC()
	_dcbae.Drawing = wml.NewCT_Drawing()
	_gabdb := wml.NewWdAnchor()
	_eaaa := AnchoredDrawing{_ffacf._dbddf, _gabdb}
	_gabdb.SimplePosAttr = unioffice.Bool(false)
	_gabdb.AllowOverlapAttr = true
	_gabdb.CNvGraphicFramePr = dml.NewCT_NonVisualGraphicFrameProperties()
	_dcbae.Drawing.Anchor = append(_dcbae.Drawing.Anchor, _gabdb)
	_gabdb.Graphic = dml.NewGraphic()
	_gabdb.Graphic.GraphicData = dml.NewCT_GraphicalObjectData()
	_gabdb.Graphic.GraphicData.UriAttr = "\u0068\u0074\u0074\u0070\u003a\u002f/\u0073\u0063\u0068e\u006d\u0061\u0073.\u006f\u0070\u0065\u006e\u0078\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002e\u006frg\u002f\u0064\u0072\u0061\u0077\u0069\u006e\u0067\u006d\u006c\u002f\u0032\u0030\u0030\u0036\u002f\u0070\u0069\u0063\u0074\u0075\u0072\u0065"
	_gabdb.SimplePos.XAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_gabdb.SimplePos.YAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_gabdb.PositionH.RelativeFromAttr = wml.WdST_RelFromHPage
	_gabdb.PositionH.Choice = &wml.WdCT_PosHChoice{}
	_gabdb.PositionH.Choice.PosOffset = unioffice.Int32(0)
	_gabdb.PositionV.RelativeFromAttr = wml.WdST_RelFromVPage
	_gabdb.PositionV.Choice = &wml.WdCT_PosVChoice{}
	_gabdb.PositionV.Choice.PosOffset = unioffice.Int32(0)
	_gabdb.Extent.CxAttr = int64(float64(img.Size().X*measurement.Pixel72) / measurement.EMU)
	_gabdb.Extent.CyAttr = int64(float64(img.Size().Y*measurement.Pixel72) / measurement.EMU)
	_gabdb.Choice = &wml.WdEG_WrapTypeChoice{}
	_gabdb.Choice.WrapSquare = wml.NewWdCT_WrapSquare()
	_gabdb.Choice.WrapSquare.WrapTextAttr = wml.WdST_WrapTextBothSides
	_edgeeb := 0x7FFFFFFF & rand.Uint32()
	_gabdb.DocPr.IdAttr = _edgeeb
	_efbad := picture.NewPic()
	_efbad.NvPicPr.CNvPr.IdAttr = _edgeeb
	_abeec := img.RelID()
	if _abeec == "" {
		return _eaaa, errors.New("\u0063\u006f\u0075\u006c\u0064\u006e\u0027\u0074\u0020\u0066\u0069\u006e\u0064\u0020\u0072\u0065\u0066\u0065\u0072\u0065n\u0063\u0065\u0020\u0074\u006f\u0020\u0069\u006d\u0061g\u0065\u0020\u0077\u0069\u0074\u0068\u0069\u006e\u0020\u0064\u006f\u0063\u0075m\u0065\u006e\u0074\u0020\u0072\u0065l\u0061\u0074\u0069o\u006e\u0073")
	}
	_gabdb.Graphic.GraphicData.Any = append(_gabdb.Graphic.GraphicData.Any, _efbad)
	_efbad.BlipFill = dml.NewCT_BlipFillProperties()
	_efbad.BlipFill.Blip = dml.NewCT_Blip()
	_efbad.BlipFill.Blip.EmbedAttr = &_abeec
	_efbad.BlipFill.Stretch = dml.NewCT_StretchInfoProperties()
	_efbad.BlipFill.Stretch.FillRect = dml.NewCT_RelativeRect()
	_efbad.SpPr = dml.NewCT_ShapeProperties()
	_efbad.SpPr.Xfrm = dml.NewCT_Transform2D()
	_efbad.SpPr.Xfrm.Off = dml.NewCT_Point2D()
	_efbad.SpPr.Xfrm.Off.XAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_efbad.SpPr.Xfrm.Off.YAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_efbad.SpPr.Xfrm.Ext = dml.NewCT_PositiveSize2D()
	_efbad.SpPr.Xfrm.Ext.CxAttr = int64(img.Size().X * measurement.Point)
	_efbad.SpPr.Xfrm.Ext.CyAttr = int64(img.Size().Y * measurement.Point)
	_efbad.SpPr.PrstGeom = dml.NewCT_PresetGeometry2D()
	_efbad.SpPr.PrstGeom.PrstAttr = dml.ST_ShapeTypeRect
	return _eaaa, nil
}

// SetShapeStyle sets style to the element v:shape in watermark.
func (_ggace *WatermarkPicture) SetShapeStyle(shapeStyle vmldrawing.ShapeStyle) {
	if _ggace._fdgfa != nil {
		_acebg := shapeStyle.String()
		_ggace._fdgfa.StyleAttr = &_acebg
	}
}

// Text return node and its child text,
func (_edfd *Node) Text() string {
	_gbee := bytes.NewBuffer([]byte{})
	switch _aabg := _edfd.X().(type) {
	case *Paragraph:
		for _, _cfbdd := range _aabg.Runs() {
			if _cfbdd.Text() != "" {
				_gbee.WriteString(_cfbdd.Text())
				_gbee.WriteString("\u000a")
			}
		}
	}
	for _, _bfdf := range _edfd.Children {
		_gbee.WriteString(_bfdf.Text())
	}
	return _gbee.String()
}

// X returns the inner wrapped XML type.
func (_fffa Paragraph) X() *wml.CT_P { return _fffa._eagd }
func _ecdc(_gbfg *Document, _bcab []*wml.CT_P, _beeg *TableInfo, _abcbd *DrawingInfo) []Node {
	_ggdf := []Node{}
	for _, _dbagb := range _bcab {
		_cgdc := Paragraph{_gbfg, _dbagb}
		_aeebb := Node{_cdbd: _gbfg, _ggda: &_cgdc}
		if _aeege, _baeb := _gbfg.Styles.SearchStyleById(_cgdc.Style()); _baeb {
			_aeebb.Style = _aeege
		}
		for _, _eefd := range _cgdc.Runs() {
			_aeebb.Children = append(_aeebb.Children, Node{_cdbd: _gbfg, _ggda: _eefd, AnchoredDrawings: _eefd.DrawingAnchored(), InlineDrawings: _eefd.DrawingInline()})
		}
		_ggdf = append(_ggdf, _aeebb)
	}
	return _ggdf
}

// ParagraphProperties returns the paragraph style properties.
func (_eaff Style) ParagraphProperties() ParagraphStyleProperties {
	if _eaff._gaege.PPr == nil {
		_eaff._gaege.PPr = wml.NewCT_PPrGeneral()
	}
	return ParagraphStyleProperties{_eaff._gaege.PPr}
}

// DrawingAnchored returns a slice of AnchoredDrawings.
func (_gcaabe Run) DrawingAnchored() []AnchoredDrawing {
	_affcg := []AnchoredDrawing{}
	for _, _fcgd := range _gcaabe._adaad.EG_RunInnerContent {
		if _fcgd.Drawing == nil {
			continue
		}
		for _, _fdeeb := range _fcgd.Drawing.Anchor {
			_affcg = append(_affcg, AnchoredDrawing{_gcaabe._dbddf, _fdeeb})
		}
	}
	return _affcg
}

// InitializeDefault constructs a default numbering.
func (_dccc Numbering) InitializeDefault() {
	_gaeg := wml.NewCT_AbstractNum()
	_gaeg.MultiLevelType = wml.NewCT_MultiLevelType()
	_gaeg.MultiLevelType.ValAttr = wml.ST_MultiLevelTypeHybridMultilevel
	_dccc._cbag.AbstractNum = append(_dccc._cbag.AbstractNum, _gaeg)
	_gaeg.AbstractNumIdAttr = 1
	const _cded = 720
	const _dfec = 720
	const _acdef = 360
	for _cbfe := 0; _cbfe < 9; _cbfe++ {
		_ggcef := wml.NewCT_Lvl()
		_ggcef.IlvlAttr = int64(_cbfe)
		_ggcef.Start = wml.NewCT_DecimalNumber()
		_ggcef.Start.ValAttr = 1
		_ggcef.NumFmt = wml.NewCT_NumFmt()
		_ggcef.NumFmt.ValAttr = wml.ST_NumberFormatBullet
		_ggcef.Suff = wml.NewCT_LevelSuffix()
		_ggcef.Suff.ValAttr = wml.ST_LevelSuffixNothing
		_ggcef.LvlText = wml.NewCT_LevelText()
		_ggcef.LvlText.ValAttr = unioffice.String("\uf0b7")
		_ggcef.LvlJc = wml.NewCT_Jc()
		_ggcef.LvlJc.ValAttr = wml.ST_JcLeft
		_ggcef.RPr = wml.NewCT_RPr()
		_ggcef.RPr.RFonts = wml.NewCT_Fonts()
		_ggcef.RPr.RFonts.AsciiAttr = unioffice.String("\u0053\u0079\u006d\u0062\u006f\u006c")
		_ggcef.RPr.RFonts.HAnsiAttr = unioffice.String("\u0053\u0079\u006d\u0062\u006f\u006c")
		_ggcef.RPr.RFonts.HintAttr = wml.ST_HintDefault
		_ggcef.PPr = wml.NewCT_PPrGeneral()
		_bcgg := int64(_cbfe*_dfec + _cded)
		_ggcef.PPr.Ind = wml.NewCT_Ind()
		_ggcef.PPr.Ind.LeftAttr = &wml.ST_SignedTwipsMeasure{}
		_ggcef.PPr.Ind.LeftAttr.Int64 = unioffice.Int64(_bcgg)
		_ggcef.PPr.Ind.HangingAttr = &sharedTypes.ST_TwipsMeasure{}
		_ggcef.PPr.Ind.HangingAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(_acdef))
		_gaeg.Lvl = append(_gaeg.Lvl, _ggcef)
	}
	_bfbee := wml.NewCT_Num()
	_bfbee.NumIdAttr = 1
	_bfbee.AbstractNumId = wml.NewCT_DecimalNumber()
	_bfbee.AbstractNumId.ValAttr = 1
	_dccc._cbag.Num = append(_dccc._cbag.Num, _bfbee)
}

// Style return the table style.
func (tbl Table) Style() string {
	if tbl.ctTbl.TblPr != nil && tbl.ctTbl.TblPr.TblStyle != nil {
		return tbl.ctTbl.TblPr.TblStyle.ValAttr
	}
	return ""
}

// WatermarkPicture is watermark picture within document.
type WatermarkPicture struct {
	_cdff  *wml.CT_Picture
	_gfedc *vmldrawing.ShapeStyle
	_fdgfa *vml.Shape
	_acbd  *vml.Shapetype
}

func _acgc(_bdgad *dml.CT_Blip, _gbaa map[string]string) {
	if _bdgad.EmbedAttr != nil {
		if _ggce, _ggcea := _gbaa[*_bdgad.EmbedAttr]; _ggcea {
			*_bdgad.EmbedAttr = _ggce
		}
	}
}

// AddSection adds a new document section with an optional section break.  If t
// is ST_SectionMarkUnset, then no break will be inserted.
func (_caacg ParagraphProperties) AddSection(t wml.ST_SectionMark) Section {
	_caacg._dfaf.SectPr = wml.NewCT_SectPr()
	if t != wml.ST_SectionMarkUnset {
		_caacg._dfaf.SectPr.Type = wml.NewCT_SectType()
		_caacg._dfaf.SectPr.Type.ValAttr = t
	}
	return Section{_caacg._aage, _caacg._dfaf.SectPr}
}

// SetImprint sets the run to imprinted text.
func (_adfac RunProperties) SetImprint(b bool) {
	if !b {
		_adfac._gbdb.Imprint = nil
	} else {
		_adfac._gbdb.Imprint = wml.NewCT_OnOff()
	}
}

// UnderlineColor returns the hex color value of paragraph underline.
func (_eccf ParagraphProperties) UnderlineColor() string {
	if _cggb := _eccf._dfaf.RPr.U; _cggb != nil {
		_dceg := _cggb.ColorAttr
		if _dceg != nil && _dceg.ST_HexColorRGB != nil {
			return *_dceg.ST_HexColorRGB
		}
	}
	return ""
}

// Section return paragraph properties section value.
func (_ebaba ParagraphProperties) Section() (Section, bool) {
	if _ebaba._dfaf.SectPr != nil {
		return Section{_ebaba._aage, _ebaba._dfaf.SectPr}, true
	}
	return Section{}, false
}

// Paragraphs returns all of the paragraphs in the document body including tables.
func (_adaf *Document) Paragraphs() []Paragraph {
	_gbg := []Paragraph{}
	if _adaf.doc.Body == nil {
		return nil
	}
	for _, _bgb := range _adaf.doc.Body.EG_BlockLevelElts {
		for _, _cge := range _bgb.EG_ContentBlockContent {
			for _, _ccbb := range _cge.P {
				_gbg = append(_gbg, Paragraph{_adaf, _ccbb})
			}
		}
	}
	for _, _cffa := range _adaf.Tables() {
		for _, _dcbf := range _cffa.Rows() {
			for _, _dbgf := range _dcbf.Cells() {
				_gbg = append(_gbg, _dbgf.Paragraphs()...)
			}
		}
	}
	return _gbg
}

// Style is a style within the styles.xml file.
type Style struct{ _gaege *wml.CT_Style }

// Emboss returns true if run emboss is on.
func (_dcfcg RunProperties) Emboss() bool { return _cadf(_dcfcg._gbdb.Emboss) }
func _eebg(_bedb *wml.CT_Tbl, _dagfe, _fgecf map[int64]int64) {
	for _, _edbg := range _bedb.EG_ContentRowContent {
		for _, _ffac := range _edbg.Tr {
			for _, _edag := range _ffac.EG_ContentCellContent {
				for _, _cde := range _edag.Tc {
					for _, _egcf := range _cde.EG_BlockLevelElts {
						for _, _adde := range _egcf.EG_ContentBlockContent {
							for _, _eegg := range _adde.P {
								_bfgff(_eegg, _dagfe, _fgecf)
							}
							for _, _bgdg := range _adde.Tbl {
								_eebg(_bgdg, _dagfe, _fgecf)
							}
						}
					}
				}
			}
		}
	}
}

// CharacterSpacingValue returns the value of run's characters spacing in twips (1/20 of point).
func (_dgfec RunProperties) CharacterSpacingValue() int64 {
	if _beec := _dgfec._gbdb.Spacing; _beec != nil {
		_gdeb := _beec.ValAttr
		if _gdeb.Int64 != nil {
			return *_gdeb.Int64
		}
	}
	return int64(0)
}

// RightToLeft returns true if run text goes from right to left.
func (_bcfg RunProperties) RightToLeft() bool { return _cadf(_bcfg._gbdb.Rtl) }

// DrawingInline return a slice of InlineDrawings.
func (_gbde Run) DrawingInline() []InlineDrawing {
	_eced := []InlineDrawing{}
	for _, _dddc := range _gbde._adaad.EG_RunInnerContent {
		if _dddc.Drawing == nil {
			continue
		}
		for _, _cafb := range _dddc.Drawing.Inline {
			_eced = append(_eced, InlineDrawing{_gbde._dbddf, _cafb})
		}
	}
	return _eced
}

// SetBottom sets the bottom border to a specified type, color and thickness.
func (_gfd CellBorders) SetBottom(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_gfd._gf.Bottom = wml.NewCT_Border()
	_feadc(_gfd._gf.Bottom, t, c, thickness)
}

// RemoveParagraph removes a paragraph from a footer.
func (_fafb Header) RemoveParagraph(p Paragraph) {
	for _, _cdbe := range _fafb._deae.EG_ContentBlockContent {
		for _babf, _aaage := range _cdbe.P {
			if _aaage == p._eagd {
				copy(_cdbe.P[_babf:], _cdbe.P[_babf+1:])
				_cdbe.P = _cdbe.P[0 : len(_cdbe.P)-1]
				return
			}
		}
	}
}

// Underline returns the type of paragraph underline.
func (_ddee ParagraphProperties) Underline() wml.ST_Underline {
	if _gfebb := _ddee._dfaf.RPr.U; _gfebb != nil {
		return _gfebb.ValAttr
	}
	return 0
}
func (_afgdg Run) newIC() *wml.EG_RunInnerContent {
	_agcdg := wml.NewEG_RunInnerContent()
	_afgdg._adaad.EG_RunInnerContent = append(_afgdg._adaad.EG_RunInnerContent, _agcdg)
	return _agcdg
}

// InsertStyle insert style to styles.
func (_geefg Styles) InsertStyle(ss Style) { _geefg._abca.Style = append(_geefg._abca.Style, ss.X()) }

// Nodes return the document's element as nodes.
func (_gdgc *Document) Nodes() Nodes {
	_ffgf := []Node{}
	for _, _cgfe := range _gdgc.doc.Body.EG_BlockLevelElts {
		_ffgf = append(_ffgf, _beaea(_gdgc, _cgfe.EG_ContentBlockContent, nil)...)
	}
	if _gdgc.doc.Body.SectPr != nil {
		_ffgf = append(_ffgf, Node{_ggda: _gdgc.doc.Body.SectPr})
	}
	_agecf := Nodes{_gabfc: _ffgf}
	return _agecf
}

// Style returns the style for a paragraph, or an empty string if it is unset.
func (_cgdeg ParagraphProperties) Style() string {
	if _cgdeg._dfaf.PStyle != nil {
		return _cgdeg._dfaf.PStyle.ValAttr
	}
	return ""
}

// SetFirstLineIndent controls the indentation of the first line in a paragraph.
func (_gcbab Paragraph) SetFirstLineIndent(m measurement.Distance) {
	_gcbab.ensurePPr()
	_fedg := _gcbab._eagd.PPr
	if _fedg.Ind == nil {
		_fedg.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_fedg.Ind.FirstLineAttr = nil
	} else {
		_fedg.Ind.FirstLineAttr = &sharedTypes.ST_TwipsMeasure{}
		_fedg.Ind.FirstLineAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

// Name returns the name of the bookmark whcih is the document unique ID that
// identifies the bookmark.
func (_cbc Bookmark) Name() string { return _cbc._gc.NameAttr }
func (_aedb *WatermarkText) getShape() *unioffice.XSDAny {
	return _aedb.getInnerElement("\u0073\u0068\u0061p\u0065")
}

// Row is a row within a table within a document.
type Row struct {
	doc *Document
	ctRow  *wml.CT_Row
}

func _bfgff(_cegeg *wml.CT_P, _gde, _fced map[int64]int64) {
	for _, _cffg := range _cegeg.EG_PContent {
		for _, _beee := range _cffg.EG_ContentRunContent {
			if _beee.R != nil {
				for _, _fabg := range _beee.R.EG_RunInnerContent {
					_edbe := _fabg.EndnoteReference
					if _edbe != nil && _edbe.IdAttr > 0 {
						if _cfed, _ffgb := _fced[_edbe.IdAttr]; _ffgb {
							_edbe.IdAttr = _cfed
						}
					}
					_beaeg := _fabg.FootnoteReference
					if _beaeg != nil && _beaeg.IdAttr > 0 {
						if _ceaa, _abbcg := _gde[_beaeg.IdAttr]; _abbcg {
							_beaeg.IdAttr = _ceaa
						}
					}
				}
			}
		}
	}
}
func _agefe() *vml.Textpath {
	_bbdeg := vml.NewTextpath()
	_bbdeg.OnAttr = sharedTypes.ST_TrueFalseTrue
	_bbdeg.FitshapeAttr = sharedTypes.ST_TrueFalseTrue
	return _bbdeg
}

// Fonts allows manipulating a style or run's fonts.
type Fonts struct{ _feae *wml.CT_Fonts }

// Properties returns the numbering level paragraph properties.
func (nl NumberingLevel) Properties() ParagraphStyleProperties {
	if nl.lvl.PPr == nil {
		nl.lvl.PPr = wml.NewCT_PPrGeneral()
	}
	return ParagraphStyleProperties{nl.lvl.PPr}
}

// Footnotes returns the footnotes defined in the document.
func (_ddag *Document) Footnotes() []Footnote {
	_faf := []Footnote{}
	for _, _bdea := range _ddag._beg.CT_Footnotes.Footnote {
		_faf = append(_faf, Footnote{_ddag, _bdea})
	}
	return _faf
}

// GetShapeStyle returns string style of the shape in watermark and format it to ShapeStyle.
func (_cbea *WatermarkPicture) GetShapeStyle() vmldrawing.ShapeStyle {
	if _cbea._fdgfa != nil && _cbea._fdgfa.StyleAttr != nil {
		return vmldrawing.NewShapeStyle(*_cbea._fdgfa.StyleAttr)
	}
	return vmldrawing.NewShapeStyle("")
}

// Outline returns true if paragraph outline is on.
func (_cadgg ParagraphProperties) Outline() bool { return _cadf(_cadgg._dfaf.RPr.Outline) }

// SetHangingIndent controls the indentation of the non-first lines in a paragraph.
func (_ffbdg ParagraphProperties) SetHangingIndent(m measurement.Distance) {
	if _ffbdg._dfaf.Ind == nil {
		_ffbdg._dfaf.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_ffbdg._dfaf.Ind.HangingAttr = nil
	} else {
		_ffbdg._dfaf.Ind.HangingAttr = &sharedTypes.ST_TwipsMeasure{}
		_ffbdg._dfaf.Ind.HangingAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

// Pict returns the pict object.
func (_cbgg *WatermarkText) Pict() *wml.CT_Picture { return _cbgg._cegfa }
func (_eeae Paragraph) ensurePPr() {
	if _eeae._eagd.PPr == nil {
		_eeae._eagd.PPr = wml.NewCT_PPr()
	}
}
func _gcad() *vml.OfcLock {
	_ebbff := vml.NewOfcLock()
	_ebbff.ExtAttr = vml.ST_ExtEdit
	_ebbff.TextAttr = sharedTypes.ST_TrueFalseTrue
	_ebbff.ShapetypeAttr = sharedTypes.ST_TrueFalseTrue
	return _ebbff
}

// SetVerticalAlignment controls the vertical alignment of the run, this is used
// to control if text is superscript/subscript.
func (_fedgf RunProperties) SetVerticalAlignment(v sharedTypes.ST_VerticalAlignRun) {
	if v == sharedTypes.ST_VerticalAlignRunUnset {
		_fedgf._gbdb.VertAlign = nil
	} else {
		_fedgf._gbdb.VertAlign = wml.NewCT_VerticalAlignRun()
		_fedgf._gbdb.VertAlign.ValAttr = v
	}
}

// CharacterSpacingMeasure returns paragraph characters spacing with its measure which can be mm, cm, in, pt, pc or pi.
func (_adcgf RunProperties) CharacterSpacingMeasure() string {
	if _feeee := _adcgf._gbdb.Spacing; _feeee != nil {
		_gagbb := _feeee.ValAttr
		if _gagbb.ST_UniversalMeasure != nil {
			return *_gagbb.ST_UniversalMeasure
		}
	}
	return ""
}
func _cegfac() *vml.Path {
	_bdcd := vml.NewPath()
	_bdcd.TextpathokAttr = sharedTypes.ST_TrueFalseTrue
	_bdcd.ConnecttypeAttr = vml.OfcST_ConnectTypeCustom
	_cdagd := "\u0040\u0039\u002c0;\u0040\u0031\u0030\u002c\u0031\u0030\u0038\u0030\u0030;\u00401\u0031,\u00321\u0036\u0030\u0030\u003b\u0040\u0031\u0032\u002c\u0031\u0030\u0038\u0030\u0030"
	_bdcd.ConnectlocsAttr = &_cdagd
	_gaaaa := "\u0032\u0037\u0030,\u0031\u0038\u0030\u002c\u0039\u0030\u002c\u0030"
	_bdcd.ConnectanglesAttr = &_gaaaa
	return _bdcd
}

// SetBottom sets the cell bottom margin
func (_aeeg CellMargins) SetBottom(d measurement.Distance) {
	_aeeg._cdae.Bottom = wml.NewCT_TblWidth()
	_age(_aeeg._cdae.Bottom, d)
}

// SetCellSpacingPercent sets the cell spacing within a table to a percent width.
func (_bccca TableStyleProperties) SetCellSpacingPercent(pct float64) {
	_bccca._degc.TblCellSpacing = wml.NewCT_TblWidth()
	_bccca._degc.TblCellSpacing.TypeAttr = wml.ST_TblWidthPct
	_bccca._degc.TblCellSpacing.WAttr = &wml.ST_MeasurementOrPercent{}
	_bccca._degc.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_bccca._degc.TblCellSpacing.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(pct * 50))
}

// TableWidth controls width values in table settings.
type TableWidth struct{ _egbb *wml.CT_TblWidth }

func (_ccge Paragraph) addStartBookmark(_adcb int64, _cffc string) *wml.CT_Bookmark {
	_gddc := wml.NewEG_PContent()
	_ccge._eagd.EG_PContent = append(_ccge._eagd.EG_PContent, _gddc)
	_cccf := wml.NewEG_ContentRunContent()
	_eabga := wml.NewEG_RunLevelElts()
	_adad := wml.NewEG_RangeMarkupElements()
	_aebd := wml.NewCT_Bookmark()
	_aebd.NameAttr = _cffc
	_aebd.IdAttr = _adcb
	_adad.BookmarkStart = _aebd
	_gddc.EG_ContentRunContent = append(_gddc.EG_ContentRunContent, _cccf)
	_cccf.EG_RunLevelElts = append(_cccf.EG_RunLevelElts, _eabga)
	_eabga.EG_RangeMarkupElements = append(_eabga.EG_RangeMarkupElements, _adad)
	return _aebd
}

// NumberingLevel is the definition for numbering for a particular level within
// a NumberingDefinition.
type NumberingLevel struct {
	lvl *wml.CT_Lvl
}

// AddTable adds a table to the table cell.
func (_dgad Cell) AddTable() Table {
	_bdg := wml.NewEG_BlockLevelElts()
	_dgad._gge.EG_BlockLevelElts = append(_dgad._gge.EG_BlockLevelElts, _bdg)
	_bae := wml.NewEG_ContentBlockContent()
	_bdg.EG_ContentBlockContent = append(_bdg.EG_ContentBlockContent, _bae)
	_cbb := wml.NewCT_Tbl()
	_bae.Tbl = append(_bae.Tbl, _cbb)
	return Table{_dgad._dga, _cbb}
}

// IsFootnote returns a bool based on whether the run has a
// footnote or not. Returns both a bool as to whether it has
// a footnote as well as the ID of the footnote.
func (_aggc Run) IsFootnote() (bool, int64) {
	if _aggc._adaad.EG_RunInnerContent != nil {
		if _aggc._adaad.EG_RunInnerContent[0].FootnoteReference != nil {
			return true, _aggc._adaad.EG_RunInnerContent[0].FootnoteReference.IdAttr
		}
	}
	return false, 0
}

// GetFooter gets a section Footer for given type
func (_gaeac Section) GetFooter(t wml.ST_HdrFtr) (Footer, bool) {
	for _, _eebb := range _gaeac._ddcag.EG_HdrFtrReferences {
		if _eebb.FooterReference.TypeAttr == t {
			for _, _bgff := range _gaeac._afafb.Footers() {
				_dgdga := _gaeac._afafb._dab.FindRIDForN(_bgff.Index(), unioffice.FooterType)
				if _dgdga == _eebb.FooterReference.IdAttr {
					return _bgff, true
				}
			}
		}
	}
	return Footer{}, false
}

// SetAlignment sets the paragraph alignment
func (_dfce NumberingLevel) SetAlignment(j wml.ST_Jc) {
	if j == wml.ST_JcUnset {
		_dfce.lvl.LvlJc = nil
	} else {
		_dfce.lvl.LvlJc = wml.NewCT_Jc()
		_dfce.lvl.LvlJc.ValAttr = j
	}
}

// SetNumberingDefinitionByID sets the numbering definition ID directly, which must
// match an ID defined in numbering.xml
func (_bdeef Paragraph) SetNumberingDefinitionByID(abstractNumberID int64) {
	_bdeef.ensurePPr()
	if _bdeef._eagd.PPr.NumPr == nil {
		_bdeef._eagd.PPr.NumPr = wml.NewCT_NumPr()
	}
	_bfadf := wml.NewCT_DecimalNumber()
	_bfadf.ValAttr = int64(abstractNumberID)
	_bdeef._eagd.PPr.NumPr.NumId = _bfadf
}

// HasFootnotes returns a bool based on the presence or abscence of footnotes within
// the document.
func (_ageb *Document) HasFootnotes() bool { return _ageb._beg != nil }

// X returns the inner wrapped XML type.
func (_ccagg TableConditionalFormatting) X() *wml.CT_TblStylePr { return _ccagg._ecbge }

// Paragraphs returns the paragraphs defined in a header.
func (_befb Header) Paragraphs() []Paragraph {
	_effa := []Paragraph{}
	for _, _accgf := range _befb._deae.EG_ContentBlockContent {
		for _, _abdb := range _accgf.P {
			_effa = append(_effa, Paragraph{_befb._dbagd, _abdb})
		}
	}
	for _, _dcgg := range _befb.Tables() {
		for _, _gfcc := range _dcgg.Rows() {
			for _, _eadbfe := range _gfcc.Cells() {
				_effa = append(_effa, _eadbfe.Paragraphs()...)
			}
		}
	}
	return _effa
}

// ExtractText returns text from the document as a DocText object.
func (_ecee *Document) ExtractText() *DocText {
	_cfbf := []TextItem{}
	for _, _eagg := range _ecee.doc.Body.EG_BlockLevelElts {
		_cfbf = append(_cfbf, _dcbb(_eagg.EG_ContentBlockContent, nil)...)
	}
	var _ebde []listItemInfo
	_eccce := _ecee.Paragraphs()
	for _, _dadcf := range _eccce {
		_acf := _cbcf(_ecee, _dadcf)
		_ebde = append(_ebde, _acf)
	}
	_gddd := _caac(_ecee)
	return &DocText{Items: _cfbf, _aefd: _ebde, _fddc: _gddd}
}

// VerticalAlign returns the value of paragraph vertical align.
func (_ggcbd ParagraphProperties) VerticalAlignment() sharedTypes.ST_VerticalAlignRun {
	if _afcf := _ggcbd._dfaf.RPr.VertAlign; _afcf != nil {
		return _afcf.ValAttr
	}
	return 0
}

// SetTextWrapSquare sets the text wrap to square with a given wrap type.
func (_ed AnchoredDrawing) SetTextWrapSquare(t wml.WdST_WrapText) {
	_ed._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_ed._dgc.Choice.WrapSquare = wml.NewWdCT_WrapSquare()
	_ed._dgc.Choice.WrapSquare.WrapTextAttr = t
}

// Clear content of node element.
func (_dgcbf *Node) Clear() { _dgcbf._ggda = nil }

// X returns the inner wrapped XML type.
func (_cfga ParagraphStyleProperties) X() *wml.CT_PPrGeneral { return _cfga._gfee }

// ReplaceTextByRegexp replace the text within node using regexp expression.
func (_dfffe *Node) ReplaceTextByRegexp(rgx *regexp.Regexp, newText string) {
	switch _fbgc := _dfffe.X().(type) {
	case *Paragraph:
		for _, _caeg := range _fbgc.Runs() {
			for _, _bbfd := range _caeg._adaad.EG_RunInnerContent {
				if _bbfd.T != nil {
					_faee := _bbfd.T.Content
					_faee = rgx.ReplaceAllString(_faee, newText)
					_bbfd.T.Content = _faee
				}
			}
		}
	}
	for _, _gbef := range _dfffe.Children {
		_gbef.ReplaceTextByRegexp(rgx, newText)
	}
}

// SetBottom sets the bottom border to a specified type, color and thickness.
func (_bacc TableBorders) SetBottom(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_bacc._gcdf.Bottom = wml.NewCT_Border()
	_feadc(_bacc._gcdf.Bottom, t, c, thickness)
}

// SetWidthAuto sets the the cell width to automatic.
func (_bdf CellProperties) SetWidthAuto() {
	_bdf._cgc.TcW = wml.NewCT_TblWidth()
	_bdf._cgc.TcW.TypeAttr = wml.ST_TblWidthAuto
}

// Footer is a footer for a document section.
type Footer struct {
	_aegg *Document
	_fcc  *wml.Ftr
}

// Italic returns true if paragraph font is italic.
func (_degbfa ParagraphProperties) Italic() bool {
	_fgfa := _degbfa._dfaf.RPr
	return _cadf(_fgfa.I) || _cadf(_fgfa.ICs)
}

// AddPageBreak adds a page break to a run.
func (_fcaf Run) AddPageBreak() {
	_eggee := _fcaf.newIC()
	_eggee.Br = wml.NewCT_Br()
	_eggee.Br.TypeAttr = wml.ST_BrTypePage
}

// AbstractNumberID returns the ID that is unique within all numbering
// definitions that is used to assign the definition to a paragraph.
func (_feacc NumberingDefinition) AbstractNumberID() int64 { return _feacc._agff.AbstractNumIdAttr }

// SetBeforeAuto controls if spacing before a paragraph is automatically determined.
func (_eecb ParagraphSpacing) SetBeforeAuto(b bool) {
	if b {
		_eecb._ffede.BeforeAutospacingAttr = &sharedTypes.ST_OnOff{}
		_eecb._ffede.BeforeAutospacingAttr.Bool = unioffice.Bool(true)
	} else {
		_eecb._ffede.BeforeAutospacingAttr = nil
	}
}

// ReplaceText replace text inside node.
func (_dbfcf *Nodes) ReplaceText(oldText, newText string) {
	for _, _cecb := range _dbfcf._gabfc {
		_cecb.ReplaceText(oldText, newText)
	}
}

// Fonts returns the style's Fonts.
func (_eccgdc RunProperties) Fonts() Fonts {
	if _eccgdc._gbdb.RFonts == nil {
		_eccgdc._gbdb.RFonts = wml.NewCT_Fonts()
	}
	return Fonts{_eccgdc._gbdb.RFonts}
}

// TableInfo is used for keep information about a table, a row and a cell where the text is located.
type TableInfo struct {
	Table    *wml.CT_Tbl
	Row      *wml.CT_Row
	Cell     *wml.CT_Tc
	RowIndex int
	ColIndex int
}

// CellProperties returns the cell properties.
func (_cfeda TableConditionalFormatting) CellProperties() CellProperties {
	if _cfeda._ecbge.TcPr == nil {
		_cfeda._ecbge.TcPr = wml.NewCT_TcPr()
	}
	return CellProperties{_cfeda._ecbge.TcPr}
}
func _aefc(_efdd *wml.CT_Tbl, _bedgf map[string]string) {
	for _, _agdgg := range _efdd.EG_ContentRowContent {
		for _, _deaa := range _agdgg.Tr {
			for _, _bdee := range _deaa.EG_ContentCellContent {
				for _, _abab := range _bdee.Tc {
					for _, _cfaa := range _abab.EG_BlockLevelElts {
						for _, _deda := range _cfaa.EG_ContentBlockContent {
							for _, _fad := range _deda.P {
								_bfgge(_fad, _bedgf)
							}
							for _, _aegb := range _deda.Tbl {
								_aefc(_aegb, _bedgf)
							}
						}
					}
				}
			}
		}
	}
}

// Text returns text from the document as one string separated with line breaks.
func (_egf *DocText) Text() string {
	_efad := bytes.NewBuffer([]byte{})
	for _, _aaae := range _egf.Items {
		if _aaae.Text != "" {
			_efad.WriteString(_aaae.Text)
			_efad.WriteString("\u000a")
		}
	}
	return _efad.String()
}

// ComplexSizeValue returns the value of paragraph font size for complex fonts in points.
func (_cfce ParagraphProperties) ComplexSizeValue() float64 {
	if _cfca := _cfce._dfaf.RPr.SzCs; _cfca != nil {
		_cedee := _cfca.ValAttr
		if _cedee.ST_UnsignedDecimalNumber != nil {
			return float64(*_cedee.ST_UnsignedDecimalNumber) / 2
		}
	}
	return 0.0
}

// ComplexSizeMeasure returns font with its measure which can be mm, cm, in, pt, pc or pi.
func (_fedf RunProperties) ComplexSizeMeasure() string {
	if _ggdcg := _fedf._gbdb.SzCs; _ggdcg != nil {
		_gaggg := _ggdcg.ValAttr
		if _gaggg.ST_PositiveUniversalMeasure != nil {
			return *_gaggg.ST_PositiveUniversalMeasure
		}
	}
	return ""
}

// SetRight sets the right border to a specified type, color and thickness.
func (_egaab ParagraphBorders) SetRight(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_egaab._fdge.Right = wml.NewCT_Border()
	_bbgf(_egaab._fdge.Right, t, c, thickness)
}

// Clear clears all content within a footer
func (_dcadb Footer) Clear() { _dcadb._fcc.EG_ContentBlockContent = nil }

// SetTextStyleItalic set text style of watermark to italic.
func (_abfg *WatermarkText) SetTextStyleItalic(value bool) {
	if _abfg._bfbf != nil {
		_caaf := _abfg.GetStyle()
		_caaf.SetItalic(value)
		_abfg.SetStyle(_caaf)
	}
}

// SetWidthPercent sets the cell to a width percentage.
func (_gaa CellProperties) SetWidthPercent(pct float64) {
	_gaa._cgc.TcW = wml.NewCT_TblWidth()
	_gaa._cgc.TcW.TypeAttr = wml.ST_TblWidthPct
	_gaa._cgc.TcW.WAttr = &wml.ST_MeasurementOrPercent{}
	_gaa._cgc.TcW.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_gaa._cgc.TcW.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(pct * 50))
}

// CharacterSpacingValue returns the value of characters spacing in twips (1/20 of point).
func (_fdee ParagraphProperties) CharacterSpacingValue() int64 {
	if _dfbgbd := _fdee._dfaf.RPr.Spacing; _dfbgbd != nil {
		_deea := _dfbgbd.ValAttr
		if _deea.Int64 != nil {
			return *_deea.Int64
		}
	}
	return int64(0)
}

// SetRightIndent controls right indent of paragraph.
func (_dcge Paragraph) SetRightIndent(m measurement.Distance) {
	_dcge.ensurePPr()
	_bbefg := _dcge._eagd.PPr
	if _bbefg.Ind == nil {
		_bbefg.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_bbefg.Ind.RightAttr = nil
	} else {
		_bbefg.Ind.RightAttr = &wml.ST_SignedTwipsMeasure{}
		_bbefg.Ind.RightAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// Run is a run of text within a paragraph that shares the same formatting.
type Run struct {
	_dbddf *Document
	_adaad *wml.CT_R
}

// PutNodeAfter put node to position after relativeTo.
func (_cfge *Document) PutNodeAfter(relativeTo, node Node) { _cfge.putNode(relativeTo, node, false) }

// AddStyle adds a new empty style, if styleID is already exists, it will return the style.
func (_dgaed Styles) AddStyle(styleID string, t wml.ST_StyleType, isDefault bool) Style {
	if _dfdf, _cbadd := _dgaed.SearchStyleById(styleID); _cbadd {
		return _dfdf
	}
	_adead := wml.NewCT_Style()
	_adead.TypeAttr = t
	if isDefault {
		_adead.DefaultAttr = &sharedTypes.ST_OnOff{}
		_adead.DefaultAttr.Bool = unioffice.Bool(isDefault)
	}
	_adead.StyleIdAttr = unioffice.String(styleID)
	_dgaed._abca.Style = append(_dgaed._abca.Style, _adead)
	return Style{_adead}
}

// Styles returns all styles.
func (_becfa Styles) Styles() []Style {
	_gaeb := []Style{}
	for _, _abdc := range _becfa._abca.Style {
		_gaeb = append(_gaeb, Style{_abdc})
	}
	return _gaeb
}

// RemoveParagraph removes a paragraph from the footnote.
func (_aedfb Footnote) RemoveParagraph(p Paragraph) {
	for _, _deaf := range _aedfb.content() {
		for _baffe, _cdfg := range _deaf.P {
			if _cdfg == p._eagd {
				copy(_deaf.P[_baffe:], _deaf.P[_baffe+1:])
				_deaf.P = _deaf.P[0 : len(_deaf.P)-1]
				return
			}
		}
	}
}
func _bgcg(_abee Paragraph) *wml.CT_NumPr {
	_abee.ensurePPr()
	if _abee._eagd.PPr.NumPr == nil {
		return nil
	}
	return _abee._eagd.PPr.NumPr
}

// SetPageSizeAndOrientation sets the page size and orientation for a section.
func (_cfdbb Section) SetPageSizeAndOrientation(w, h measurement.Distance, orientation wml.ST_PageOrientation) {
	if _cfdbb._ddcag.PgSz == nil {
		_cfdbb._ddcag.PgSz = wml.NewCT_PageSz()
	}
	_cfdbb._ddcag.PgSz.OrientAttr = orientation
	if orientation == wml.ST_PageOrientationLandscape {
		_cfdbb._ddcag.PgSz.WAttr = &sharedTypes.ST_TwipsMeasure{}
		_cfdbb._ddcag.PgSz.WAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(h / measurement.Twips))
		_cfdbb._ddcag.PgSz.HAttr = &sharedTypes.ST_TwipsMeasure{}
		_cfdbb._ddcag.PgSz.HAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(w / measurement.Twips))
	} else {
		_cfdbb._ddcag.PgSz.WAttr = &sharedTypes.ST_TwipsMeasure{}
		_cfdbb._ddcag.PgSz.WAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(w / measurement.Twips))
		_cfdbb._ddcag.PgSz.HAttr = &sharedTypes.ST_TwipsMeasure{}
		_cfdbb._ddcag.PgSz.HAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(h / measurement.Twips))
	}
}

// SetLeftIndent controls the left indent of the paragraph.
func (_gdbd ParagraphStyleProperties) SetLeftIndent(m measurement.Distance) {
	if _gdbd._gfee.Ind == nil {
		_gdbd._gfee.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_gdbd._gfee.Ind.LeftAttr = nil
	} else {
		_gdbd._gfee.Ind.LeftAttr = &wml.ST_SignedTwipsMeasure{}
		_gdbd._gfee.Ind.LeftAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// FindNodeByText return node based on matched text and return a slice of node.
func (_fgadf *Nodes) FindNodeByRegexp(r *regexp.Regexp) []Node {
	_gebe := []Node{}
	for _, _gggba := range _fgadf._gabfc {
		if r.MatchString(_gggba.Text()) {
			_gebe = append(_gebe, _gggba)
		}
		_cadc := Nodes{_gabfc: _gggba.Children}
		_gebe = append(_gebe, _cadc.FindNodeByRegexp(r)...)
	}
	return _gebe
}

// AddTab adds tab to a run and can be used with the the Paragraph's tab stops.
func (_ffeb Run) AddTab() { _afbc := _ffeb.newIC(); _afbc.Tab = wml.NewCT_Empty() }

const (
	OnOffValueUnset OnOffValue = iota
	OnOffValueOff
	OnOffValueOn
)

type mergeFieldInfo struct {
	_gdfge               string
	_cgec                string
	_dfga                string
	_dbfgc               bool
	_fgaa                bool
	_bddb                bool
	_bcdae               bool
	_abdbd               Paragraph
	_bbcb, _cdcbe, _gfaf int
	_ceaf                *wml.EG_PContent
	_debdg               bool
}

// GetHeader gets a section Header for given type t [ST_HdrFtrDefault, ST_HdrFtrEven, ST_HdrFtrFirst]
func (_debf Section) GetHeader(t wml.ST_HdrFtr) (Header, bool) {
	for _, _dagfb := range _debf._ddcag.EG_HdrFtrReferences {
		if _dagfb.HeaderReference.TypeAttr == t {
			for _, _gagfd := range _debf._afafb.Headers() {
				_cffaf := _debf._afafb._dab.FindRIDForN(_gagfd.Index(), unioffice.HeaderType)
				if _cffaf == _dagfb.HeaderReference.IdAttr {
					return _gagfd, true
				}
			}
		}
	}
	return Header{}, false
}
func (_gaef *Document) insertStyleFromNode(_afbb Node) {
	if _afbb.Style.X() != nil {
		if _, _dfcb := _gaef.Styles.SearchStyleById(_afbb.Style.StyleID()); !_dfcb {
			_gaef.Styles.InsertStyle(_afbb.Style)
			_fdgc := _afbb.Style.ParagraphProperties()
			_gaef.insertNumberingFromStyleProperties(_afbb._cdbd.Numbering, _fdgc)
		}
	}
}

// NumberingDefinition defines a numbering definition for a list of pragraphs.
type NumberingDefinition struct{ _agff *wml.CT_AbstractNum }

// NumId return numbering numId that being use by style properties.
func (_gadaa ParagraphStyleProperties) NumId() int64 {
	if _gadaa._gfee.NumPr != nil {
		if _gadaa._gfee.NumPr.NumId != nil {
			return _gadaa._gfee.NumPr.NumId.ValAttr
		}
	}
	return -1
}

// GetWrapPathStart return wrapPath start value.
func (_ca AnchorDrawWrapOptions) GetWrapPathStart() *dml.CT_Point2D { return _ca._dd }

// Bookmarks returns all of the bookmarks defined in the document.
func (_dag Document) Bookmarks() []Bookmark {
	if _dag.doc.Body == nil {
		return nil
	}
	_bcga := []Bookmark{}
	for _, _begeg := range _dag.doc.Body.EG_BlockLevelElts {
		for _, _gbac := range _begeg.EG_ContentBlockContent {
			for _, _afca := range _aebg(_gbac) {
				_bcga = append(_bcga, _afca)
			}
		}
	}
	return _bcga
}

// HyperLink is a link within a document.
type HyperLink struct {
	_acbfa *Document
	_baaf  *wml.CT_Hyperlink
}

// AddWatermarkPicture adds new watermark picture to document.
func (_gbf *Document) AddWatermarkPicture(imageRef common.ImageRef) WatermarkPicture {
	var _eef []Header
	if _ggb, _ddd := _gbf.BodySection().GetHeader(wml.ST_HdrFtrDefault); _ddd {
		_eef = append(_eef, _ggb)
	}
	if _bef, _cdaf := _gbf.BodySection().GetHeader(wml.ST_HdrFtrEven); _cdaf {
		_eef = append(_eef, _bef)
	}
	if _agec, _dcabd := _gbf.BodySection().GetHeader(wml.ST_HdrFtrFirst); _dcabd {
		_eef = append(_eef, _agec)
	}
	if len(_eef) < 1 {
		_agfgg := _gbf.AddHeader()
		_gbf.BodySection().SetHeader(_agfgg, wml.ST_HdrFtrDefault)
		_eef = append(_eef, _agfgg)
	}
	var _cdfbe error
	_fgca := NewWatermarkPicture()
	for _, _gcaag := range _eef {
		imageRef, _cdfbe = _gcaag.AddImageRef(imageRef)
		if _cdfbe != nil {
			return WatermarkPicture{}
		}
		_dbfga := _gcaag.Paragraphs()
		if len(_dbfga) < 1 {
			_egce := _gcaag.AddParagraph()
			_egce.AddRun().AddText("")
		}
		for _, _afadc := range _gcaag.X().EG_ContentBlockContent {
			for _, _faba := range _afadc.P {
				for _, _egbg := range _faba.EG_PContent {
					for _, _fabag := range _egbg.EG_ContentRunContent {
						if _fabag.R == nil {
							continue
						}
						for _, _dbdbc := range _fabag.R.EG_RunInnerContent {
							_dbdbc.Pict = _fgca._cdff
							break
						}
					}
				}
			}
		}
	}
	_fgca.SetPicture(imageRef)
	return _fgca
}

var _cbfg = [...]uint8{0, 20, 37, 58, 79}

// Italic returns true if run font is italic.
func (_gffc RunProperties) Italic() bool {
	_aagba := _gffc._gbdb
	return _cadf(_aagba.I) || _cadf(_aagba.ICs)
}

// ExtractFromHeader returns text from the document header as an array of TextItems.
func ExtractFromHeader(header *wml.Hdr) []TextItem { return _dcbb(header.EG_ContentBlockContent, nil) }

// AnchoredDrawing is an absolutely positioned image within a document page.
type AnchoredDrawing struct {
	_dg  *Document
	_dgc *wml.WdAnchor
}

// SetEnabled marks a FormField as enabled or disabled.
func (_eefae FormField) SetEnabled(enabled bool) {
	_gegd := wml.NewCT_OnOff()
	_gegd.ValAttr = &sharedTypes.ST_OnOff{Bool: &enabled}
	_eefae._cbde.Enabled = []*wml.CT_OnOff{_gegd}
}

// SetRight sets the right border to a specified type, color and thickness.
func (_baf CellBorders) SetRight(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_baf._gf.Right = wml.NewCT_Border()
	_feadc(_baf._gf.Right, t, c, thickness)
}

// Numbering is the document wide numbering styles contained in numbering.xml.
type Numbering struct{ _cbag *wml.Numbering }

// Properties returns the table properties.
func (_eabd Table) Properties() TableProperties {
	if _eabd.ctTbl.TblPr == nil {
		_eabd.ctTbl.TblPr = wml.NewCT_TblPr()
	}
	return TableProperties{_eabd.ctTbl.TblPr}
}

// TableBorders allows manipulation of borders on a table.
type TableBorders struct{ _gcdf *wml.CT_TblBorders }

// SetPictureWashout set washout to watermark picture.
func (_efffb *WatermarkPicture) SetPictureWashout(isWashout bool) {
	if _efffb._fdgfa != nil {
		_adaed := _efffb._fdgfa.EG_ShapeElements
		if len(_adaed) > 0 && _adaed[0].Imagedata != nil {
			if isWashout {
				_cbfaf := "\u0031\u0039\u0036\u0036\u0031\u0066"
				_ggfa := "\u0032\u0032\u0039\u0033\u0038\u0066"
				_adaed[0].Imagedata.GainAttr = &_cbfaf
				_adaed[0].Imagedata.BlacklevelAttr = &_ggfa
			}
		}
	}
}

// RStyle returns the name of character style.
// It is defined here http://officeopenxml.com/WPstyleCharStyles.php
func (_gcae RunProperties) RStyle() string {
	if _gcae._gbdb.RStyle != nil {
		return _gcae._gbdb.RStyle.ValAttr
	}
	return ""
}

// AddHyperlink adds a hyperlink to a document. Adding the hyperlink to a document
// and setting it on a cell is more efficient than setting hyperlinks directly
// on a cell.
func (_fgbd Document) AddHyperlink(url string) common.Hyperlink { return _fgbd._dab.AddHyperlink(url) }
func (_dcec FormFieldType) String() string {
	if _dcec >= FormFieldType(len(_cbfg)-1) {
		return fmt.Sprintf("\u0046\u006f\u0072\u006d\u0046\u0069\u0065\u006c\u0064\u0054\u0079\u0070e\u0028\u0025\u0064\u0029", _dcec)
	}
	return _bffd[_cbfg[_dcec]:_cbfg[_dcec+1]]
}

// SetFirstLineIndent controls the first line indent of the paragraph.
func (_dbab ParagraphStyleProperties) SetFirstLineIndent(m measurement.Distance) {
	if _dbab._gfee.Ind == nil {
		_dbab._gfee.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_dbab._gfee.Ind.FirstLineAttr = nil
	} else {
		_dbab._gfee.Ind.FirstLineAttr = &sharedTypes.ST_TwipsMeasure{}
		_dbab._gfee.Ind.FirstLineAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

// IsEndnote returns a bool based on whether the run has a
// footnote or not. Returns both a bool as to whether it has
// a footnote as well as the ID of the footnote.
func (_ffgfd Run) IsEndnote() (bool, int64) {
	if _ffgfd._adaad.EG_RunInnerContent != nil {
		if _ffgfd._adaad.EG_RunInnerContent[0].EndnoteReference != nil {
			return true, _ffgfd._adaad.EG_RunInnerContent[0].EndnoteReference.IdAttr
		}
	}
	return false, 0
}

// SetInsideHorizontal sets the interior horizontal borders to a specified type, color and thickness.
func (_aaeb TableBorders) SetInsideHorizontal(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_aaeb._gcdf.InsideH = wml.NewCT_Border()
	_feadc(_aaeb._gcdf.InsideH, t, c, thickness)
}

// SetUnderline controls underline for a run style.
func (_feggd RunProperties) SetUnderline(style wml.ST_Underline, c color.Color) {
	if style == wml.ST_UnderlineUnset {
		_feggd._gbdb.U = nil
	} else {
		_feggd._gbdb.U = wml.NewCT_Underline()
		_feggd._gbdb.U.ColorAttr = &wml.ST_HexColor{}
		_feggd._gbdb.U.ColorAttr.ST_HexColorRGB = c.AsRGBString()
		_feggd._gbdb.U.ValAttr = style
	}
}

// SetSpacing sets the spacing that comes before and after the paragraph.
// Deprecated: See Spacing() instead which allows finer control.
func (_fgea ParagraphProperties) SetSpacing(before, after measurement.Distance) {
	if _fgea._dfaf.Spacing == nil {
		_fgea._dfaf.Spacing = wml.NewCT_Spacing()
	}
	_fgea._dfaf.Spacing.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
	_fgea._dfaf.Spacing.BeforeAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(before / measurement.Twips))
	_fgea._dfaf.Spacing.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
	_fgea._dfaf.Spacing.AfterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(after / measurement.Twips))
}

// SetFirstLineIndent controls the indentation of the first line in a paragraph.
func (_ddabf ParagraphProperties) SetFirstLineIndent(m measurement.Distance) {
	if _ddabf._dfaf.Ind == nil {
		_ddabf._dfaf.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_ddabf._dfaf.Ind.FirstLineAttr = nil
	} else {
		_ddabf._dfaf.Ind.FirstLineAttr = &sharedTypes.ST_TwipsMeasure{}
		_ddabf._dfaf.Ind.FirstLineAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

var _eece = false

// SetOutline sets the run to outlined text.
func (_gffe RunProperties) SetOutline(b bool) {
	if !b {
		_gffe._gbdb.Outline = nil
	} else {
		_gffe._gbdb.Outline = wml.NewCT_OnOff()
	}
}

// TableLook is the conditional formatting associated with a table style that
// has been assigned to a table.
type TableLook struct{ ctTblLook *wml.CT_TblLook }

// IsItalic returns true if the run has been set to italics.
func (_bebad RunProperties) IsItalic() bool { return _bebad.ItalicValue() == OnOffValueOn }

// SetShading controls the cell shading.
func (_dfc CellProperties) SetShading(shd wml.ST_Shd, foreground, fill color.Color) {
	if shd == wml.ST_ShdUnset {
		_dfc._cgc.Shd = nil
	} else {
		_dfc._cgc.Shd = wml.NewCT_Shd()
		_dfc._cgc.Shd.ValAttr = shd
		_dfc._cgc.Shd.ColorAttr = &wml.ST_HexColor{}
		if foreground.IsAuto() {
			_dfc._cgc.Shd.ColorAttr.ST_HexColorAuto = wml.ST_HexColorAutoAuto
		} else {
			_dfc._cgc.Shd.ColorAttr.ST_HexColorRGB = foreground.AsRGBString()
		}
		_dfc._cgc.Shd.FillAttr = &wml.ST_HexColor{}
		if fill.IsAuto() {
			_dfc._cgc.Shd.FillAttr.ST_HexColorAuto = wml.ST_HexColorAutoAuto
		} else {
			_dfc._cgc.Shd.FillAttr.ST_HexColorRGB = fill.AsRGBString()
		}
	}
}

// AddRun adds a run to a paragraph.
func (_cbdec Paragraph) AddRun() Run {
	_eeece := wml.NewEG_PContent()
	_cbdec._eagd.EG_PContent = append(_cbdec._eagd.EG_PContent, _eeece)
	_bfcf := wml.NewEG_ContentRunContent()
	_eeece.EG_ContentRunContent = append(_eeece.EG_ContentRunContent, _bfcf)
	_edaf := wml.NewCT_R()
	_bfcf.R = _edaf
	return Run{_cbdec._fagf, _edaf}
}

// Outline returns true if run outline is on.
func (_ebac RunProperties) Outline() bool { return _cadf(_ebac._gbdb.Outline) }

// SetTableIndent sets the Table Indent from the Leading Margin
func (_dbec TableStyleProperties) SetTableIndent(ind measurement.Distance) {
	_dbec._degc.TblInd = wml.NewCT_TblWidth()
	_dbec._degc.TblInd.TypeAttr = wml.ST_TblWidthDxa
	_dbec._degc.TblInd.WAttr = &wml.ST_MeasurementOrPercent{}
	_dbec._degc.TblInd.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_dbec._degc.TblInd.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(ind / measurement.Dxa))
}

// X returns the inner wrapped XML type.
func (_bcggd Row) X() *wml.CT_Row { return _bcggd.ctRow }

// SetBold sets the run to bold.
func (_faaceb RunProperties) SetBold(b bool) {
	if !b {
		_faaceb._gbdb.B = nil
		_faaceb._gbdb.BCs = nil
	} else {
		_faaceb._gbdb.B = wml.NewCT_OnOff()
		_faaceb._gbdb.BCs = wml.NewCT_OnOff()
	}
}

const (
	FieldCurrentPage   = "\u0050\u0041\u0047\u0045"
	FieldNumberOfPages = "\u004e\u0055\u004d\u0050\u0041\u0047\u0045\u0053"
	FieldDate          = "\u0044\u0041\u0054\u0045"
	FieldCreateDate    = "\u0043\u0052\u0045\u0041\u0054\u0045\u0044\u0041\u0054\u0045"
	FieldEditTime      = "\u0045\u0044\u0049\u0054\u0054\u0049\u004d\u0045"
	FieldPrintDate     = "\u0050R\u0049\u004e\u0054\u0044\u0041\u0054E"
	FieldSaveDate      = "\u0053\u0041\u0056\u0045\u0044\u0041\u0054\u0045"
	FieldTIme          = "\u0054\u0049\u004d\u0045"
	FieldTOC           = "\u0054\u004f\u0043"
)

// SetAlignment sets the alignment of a table within the page.
func (_fbdge TableProperties) SetAlignment(align wml.ST_JcTable) {
	if align == wml.ST_JcTableUnset {
		_fbdge._efag.Jc = nil
	} else {
		_fbdge._efag.Jc = wml.NewCT_JcTable()
		_fbdge._efag.Jc.ValAttr = align
	}
}

// Type returns the type of the field.
func (_cgba FormField) Type() FormFieldType {
	if _cgba._cbde.TextInput != nil {
		return FormFieldTypeText
	} else if _cgba._cbde.CheckBox != nil {
		return FormFieldTypeCheckBox
	} else if _cgba._cbde.DdList != nil {
		return FormFieldTypeDropDown
	}
	return FormFieldTypeUnknown
}
func _gdfd(_feec []*wml.CT_P, _cbfab *TableInfo, _aece *DrawingInfo) []TextItem {
	_geea := []TextItem{}
	for _, _aefg := range _feec {
		_geea = append(_geea, _eeebf(_aefg, nil, _cbfab, _aece, _aefg.EG_PContent)...)
	}
	return _geea
}

// SetVerticalMerge controls the vertical merging of cells.
func (_fed CellProperties) SetVerticalMerge(mergeVal wml.ST_Merge) {
	if mergeVal == wml.ST_MergeUnset {
		_fed._cgc.VMerge = nil
	} else {
		_fed._cgc.VMerge = wml.NewCT_VMerge()
		_fed._cgc.VMerge.ValAttr = mergeVal
	}
}
func (_gage *Document) getWatermarkHeaderInnerContentPictures() []*wml.CT_Picture {
	var _eadbc []*wml.CT_Picture
	for _, _bcad := range _gage.Headers() {
		for _, _abag := range _bcad.X().EG_ContentBlockContent {
			for _, _egeaa := range _abag.P {
				for _, _bbed := range _egeaa.EG_PContent {
					for _, _eccc := range _bbed.EG_ContentRunContent {
						if _eccc.R == nil {
							continue
						}
						for _, _dce := range _eccc.R.EG_RunInnerContent {
							if _dce.Pict == nil {
								continue
							}
							_aggf := false
							for _, _dgag := range _dce.Pict.Any {
								_ccfd, _egdc := _dgag.(*unioffice.XSDAny)
								if _egdc && _ccfd.XMLName.Local == "\u0073\u0068\u0061p\u0065" {
									_aggf = true
								}
							}
							if _aggf {
								_eadbc = append(_eadbc, _dce.Pict)
							}
						}
					}
				}
			}
		}
	}
	return _eadbc
}

// Properties returns the paragraph properties.
func (_eege Paragraph) Properties() ParagraphProperties {
	_eege.ensurePPr()
	return ParagraphProperties{_eege._fagf, _eege._eagd.PPr}
}

// RemoveMailMerge removes any mail merge settings
func (_dbea Settings) RemoveMailMerge() { _dbea._cdbbf.MailMerge = nil }

// Paragraphs returns the paragraphs defined in the cell.
func (_abf Cell) Paragraphs() []Paragraph {
	_dba := []Paragraph{}
	for _, _aee := range _abf._gge.EG_BlockLevelElts {
		for _, _cga := range _aee.EG_ContentBlockContent {
			for _, _afdf := range _cga.P {
				_dba = append(_dba, Paragraph{_abf._dga, _afdf})
			}
		}
	}
	return _dba
}

// Paragraphs returns the paragraphs defined in a footer.
func (_baga Footer) Paragraphs() []Paragraph {
	_feace := []Paragraph{}
	for _, _ebdc := range _baga._fcc.EG_ContentBlockContent {
		for _, _dgba := range _ebdc.P {
			_feace = append(_feace, Paragraph{_baga._aegg, _dgba})
		}
	}
	for _, _bgbg := range _baga.Tables() {
		for _, _eacg := range _bgbg.Rows() {
			for _, _ceag := range _eacg.Cells() {
				_feace = append(_feace, _ceag.Paragraphs()...)
			}
		}
	}
	return _feace
}

// SetTop sets the cell top margin
func (_gbc CellMargins) SetTop(d measurement.Distance) {
	_gbc._cdae.Top = wml.NewCT_TblWidth()
	_age(_gbc._cdae.Top, d)
}
func _age(_ddb *wml.CT_TblWidth, _dad measurement.Distance) {
	_ddb.TypeAttr = wml.ST_TblWidthDxa
	_ddb.WAttr = &wml.ST_MeasurementOrPercent{}
	_ddb.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_ddb.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(_dad / measurement.Dxa))
}

// X returns the inner wrapped XML type.
func (_bfcc Endnote) X() *wml.CT_FtnEdn { return _bfcc._fagg }

// SetAll sets all of the borders to a given value.
func (_gbdf ParagraphBorders) SetAll(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_gbdf.SetBottom(t, c, thickness)
	_gbdf.SetLeft(t, c, thickness)
	_gbdf.SetRight(t, c, thickness)
	_gbdf.SetTop(t, c, thickness)
}

// SetTextWrapTopAndBottom sets the text wrap to top and bottom.
func (_cc AnchoredDrawing) SetTextWrapTopAndBottom() {
	_cc._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_cc._dgc.Choice.WrapTopAndBottom = wml.NewWdCT_WrapTopBottom()
	_cc._dgc.LayoutInCellAttr = true
	_cc._dgc.AllowOverlapAttr = true
}

// SetConformance sets conformance attribute of the document
// as one of these values from github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes:
// ST_ConformanceClassUnset, ST_ConformanceClassStrict or ST_ConformanceClassTransitional.
func (_cagb Document) SetConformance(conformanceAttr sharedTypes.ST_ConformanceClass) {
	_cagb.doc.ConformanceAttr = conformanceAttr
}

// AddCheckBox adds checkbox form field to the paragraph and returns it.
func (_fbdcad Paragraph) AddCheckBox(name string) FormField {
	_edga := _fbdcad.addFldCharsForField(name, "\u0046\u004f\u0052M\u0043\u0048\u0045\u0043\u004b\u0042\u004f\u0058")
	_edga._cbde.CheckBox = wml.NewCT_FFCheckBox()
	return _edga
}

// SetColumnSpan sets the number of Grid Columns Spanned by the Cell.  This is used
// to give the appearance of merged cells.
func (_gfe CellProperties) SetColumnSpan(cols int) {
	if cols == 0 {
		_gfe._cgc.GridSpan = nil
	} else {
		_gfe._cgc.GridSpan = wml.NewCT_DecimalNumber()
		_gfe._cgc.GridSpan.ValAttr = int64(cols)
	}
}

// Numbering return numbering that being use by paragraph.
func (_ffbd Paragraph) Numbering() Numbering {
	_ffbd.ensurePPr()
	_aacd := NewNumbering()
	if _ffbd._eagd.PPr.NumPr != nil {
		_ccce := int64(-1)
		_efbeg := int64(-1)
		if _ffbd._eagd.PPr.NumPr.NumId != nil {
			_ccce = _ffbd._eagd.PPr.NumPr.NumId.ValAttr
		}
		for _, _edbc := range _ffbd._fagf.Numbering._cbag.Num {
			if _ccce < 0 {
				break
			}
			if _edbc.NumIdAttr == _ccce {
				if _edbc.AbstractNumId != nil {
					_efbeg = _edbc.AbstractNumId.ValAttr
					_aacd._cbag.Num = append(_aacd._cbag.Num, _edbc)
					break
				}
			}
		}
		for _, _dbcaa := range _ffbd._fagf.Numbering._cbag.AbstractNum {
			if _efbeg < 0 {
				break
			}
			if _dbcaa.AbstractNumIdAttr == _efbeg {
				_aacd._cbag.AbstractNum = append(_aacd._cbag.AbstractNum, _dbcaa)
				break
			}
		}
	}
	return _aacd
}
func (_beaec Paragraph) addBeginFldChar(_dgec string) *wml.CT_FFData {
	_dggf := _beaec.addFldChar()
	_dggf.FldCharTypeAttr = wml.ST_FldCharTypeBegin
	_dggf.FfData = wml.NewCT_FFData()
	_cgfef := wml.NewCT_FFName()
	_cgfef.ValAttr = &_dgec
	_dggf.FfData.Name = []*wml.CT_FFName{_cgfef}
	return _dggf.FfData
}

// SetCellSpacingAuto sets the cell spacing within a table to automatic.
func (_bdac TableProperties) SetCellSpacingAuto() {
	_bdac._efag.TblCellSpacing = wml.NewCT_TblWidth()
	_bdac._efag.TblCellSpacing.TypeAttr = wml.ST_TblWidthAuto
}

// AddTable adds a new table to the document body.
func (_cdd *Document) AddTable() Table {
	_ded := wml.NewEG_BlockLevelElts()
	_cdd.doc.Body.EG_BlockLevelElts = append(_cdd.doc.Body.EG_BlockLevelElts, _ded)
	_ggf := wml.NewEG_ContentBlockContent()
	_ded.EG_ContentBlockContent = append(_ded.EG_ContentBlockContent, _ggf)
	_baff := wml.NewCT_Tbl()
	_ggf.Tbl = append(_ggf.Tbl, _baff)
	return Table{_cdd, _baff}
}

const _bffd = "\u0046\u006f\u0072\u006d\u0046\u0069\u0065l\u0064\u0054\u0079\u0070\u0065\u0055\u006e\u006b\u006e\u006f\u0077\u006e\u0046\u006fr\u006dF\u0069\u0065\u006c\u0064\u0054\u0079p\u0065\u0054\u0065\u0078\u0074\u0046\u006fr\u006d\u0046\u0069\u0065\u006c\u0064\u0054\u0079\u0070\u0065\u0043\u0068\u0065\u0063\u006b\u0042\u006f\u0078\u0046\u006f\u0072\u006d\u0046i\u0065\u006c\u0064\u0054\u0079\u0070\u0065\u0044\u0072\u006f\u0070\u0044\u006fw\u006e"

// Color returns the style's Color.
func (_fcbbf RunProperties) Color() Color {
	if _fcbbf._gbdb.Color == nil {
		_fcbbf._gbdb.Color = wml.NewCT_Color()
	}
	return Color{_fcbbf._gbdb.Color}
}
func _fgccb(_adeg *wml.CT_OnOff) OnOffValue {
	if _adeg == nil {
		return OnOffValueUnset
	}
	if _adeg.ValAttr != nil && _adeg.ValAttr.Bool != nil && *_adeg.ValAttr.Bool == false {
		return OnOffValueOff
	}
	return OnOffValueOn
}

// SetTargetByRef sets the URL target of the hyperlink and is more efficient if a link
// destination will be used many times.
func (_acfg HyperLink) SetTargetByRef(link common.Hyperlink) {
	_acfg._baaf.IdAttr = unioffice.String(common.Relationship(link).ID())
	_acfg._baaf.AnchorAttr = nil
}
func _cadf(_ecb *wml.CT_OnOff) bool { return _ecb != nil }

// Borders returns the ParagraphBorders for setting-up border on paragraph.
func (_agfa Paragraph) Borders() ParagraphBorders {
	_agfa.ensurePPr()
	if _agfa._eagd.PPr.PBdr == nil {
		_agfa._eagd.PPr.PBdr = wml.NewCT_PBdr()
	}
	return ParagraphBorders{_agfa._fagf, _agfa._eagd.PPr.PBdr}
}

// RunProperties returns the run style properties.
func (_efdb Style) RunProperties() RunProperties {
	if _efdb._gaege.RPr == nil {
		_efdb._gaege.RPr = wml.NewCT_RPr()
	}
	return RunProperties{_efdb._gaege.RPr}
}

// SetPrimaryStyle marks the style as a primary style.
func (_cfgd Style) SetPrimaryStyle(b bool) {
	if b {
		_cfgd._gaege.QFormat = wml.NewCT_OnOff()
	} else {
		_cfgd._gaege.QFormat = nil
	}
}

// X returns the inner wml.CT_TblBorders
func (_feade TableBorders) X() *wml.CT_TblBorders { return _feade._gcdf }

// SetStrict is a shortcut for document.SetConformance,
// as one of these values from github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes:
// ST_ConformanceClassUnset, ST_ConformanceClassStrict or ST_ConformanceClassTransitional.
func (_dcgb Document) SetStrict(strict bool) {
	if strict {
		_dcgb.doc.ConformanceAttr = sharedTypes.ST_ConformanceClassStrict
	} else {
		_dcgb.doc.ConformanceAttr = sharedTypes.ST_ConformanceClassTransitional
	}
}

// Levels returns all of the numbering levels defined in the definition.
func (_cabb NumberingDefinition) Levels() []NumberingLevel {
	_aceb := []NumberingLevel{}
	for _, _cccb := range _cabb._agff.Lvl {
		_aceb = append(_aceb, NumberingLevel{_cccb})
	}
	return _aceb
}

// SetShadow sets the run to shadowed text.
func (_deba RunProperties) SetShadow(b bool) {
	if !b {
		_deba._gbdb.Shadow = nil
	} else {
		_deba._gbdb.Shadow = wml.NewCT_OnOff()
	}
}

// SetStrikeThrough sets the run to strike-through.
func (_dfefc RunProperties) SetStrikeThrough(b bool) {
	if !b {
		_dfefc._gbdb.Strike = nil
	} else {
		_dfefc._gbdb.Strike = wml.NewCT_OnOff()
	}
}

// SetTop sets the top border to a specified type, color and thickness.
func (_bggfd TableBorders) SetTop(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_bggfd._gcdf.Top = wml.NewCT_Border()
	_feadc(_bggfd._gcdf.Top, t, c, thickness)
}

// SetWidthPercent sets the table to a width percentage.
func (_cgdgbc TableProperties) SetWidthPercent(pct float64) {
	_cgdgbc._efag.TblW = wml.NewCT_TblWidth()
	_cgdgbc._efag.TblW.TypeAttr = wml.ST_TblWidthPct
	_cgdgbc._efag.TblW.WAttr = &wml.ST_MeasurementOrPercent{}
	_cgdgbc._efag.TblW.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_cgdgbc._efag.TblW.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(pct * 50))
}

// X returns the inner wrapped XML type.
func (_aged Settings) X() *wml.Settings { return _aged._cdbbf }

// X returns the inner wrapped XML type.
func (_bffec Style) X() *wml.CT_Style { return _bffec._gaege }

// X returns the inner wml.CT_PBdr
func (_ebaa ParagraphBorders) X() *wml.CT_PBdr { return _ebaa._fdge }

// SetUISortOrder controls the order the style is displayed in the UI.
func (_bcfeg Style) SetUISortOrder(order int) {
	_bcfeg._gaege.UiPriority = wml.NewCT_DecimalNumber()
	_bcfeg._gaege.UiPriority.ValAttr = int64(order)
}

// Name returns the name of the style if set.
func (_gbbeg Style) Name() string {
	if _gbbeg._gaege.Name == nil {
		return ""
	}
	return _gbbeg._gaege.Name.ValAttr
}
func (_bgfda *WatermarkText) getInnerElement(_feba string) *unioffice.XSDAny {
	for _, _bfcfc := range _bgfda._cegfa.Any {
		_daeb, _caccg := _bfcfc.(*unioffice.XSDAny)
		if _caccg && (_daeb.XMLName.Local == _feba || _daeb.XMLName.Local == "\u0076\u003a"+_feba) {
			return _daeb
		}
	}
	return nil
}
func (_gged *Document) tables(_dcab *wml.EG_ContentBlockContent) []Table {
	_efd := []Table{}
	for _, _cgaf := range _dcab.Tbl {
		_efd = append(_efd, Table{_gged, _cgaf})
		for _, _edd := range _cgaf.EG_ContentRowContent {
			for _, _cag := range _edd.Tr {
				for _, _fcf := range _cag.EG_ContentCellContent {
					for _, _bege := range _fcf.Tc {
						for _, _gfda := range _bege.EG_BlockLevelElts {
							for _, _ebe := range _gfda.EG_ContentBlockContent {
								for _, _ffga := range _gged.tables(_ebe) {
									_efd = append(_efd, _ffga)
								}
							}
						}
					}
				}
			}
		}
	}
	return _efd
}

// Text returns the underlying tet in the run.
func (_afgfg Run) Text() string {
	if len(_afgfg._adaad.EG_RunInnerContent) == 0 {
		return ""
	}
	_eeac := bytes.Buffer{}
	for _, _adeab := range _afgfg._adaad.EG_RunInnerContent {
		if _adeab.T != nil {
			_eeac.WriteString(_adeab.T.Content)
		}
		if _adeab.Tab != nil {
			_eeac.WriteByte('\t')
		}
	}
	return _eeac.String()
}

// Style returns the style for a paragraph, or an empty string if it is unset.
func (_eabcb Paragraph) Style() string {
	if _eabcb._eagd.PPr != nil && _eabcb._eagd.PPr.PStyle != nil {
		return _eabcb._eagd.PPr.PStyle.ValAttr
	}
	return ""
}

// SetUnhideWhenUsed controls if a semi hidden style becomes visible when used.
func (_gffga Style) SetUnhideWhenUsed(b bool) {
	if b {
		_gffga._gaege.UnhideWhenUsed = wml.NewCT_OnOff()
	} else {
		_gffga._gaege.UnhideWhenUsed = nil
	}
}
func _bgbf() *vml.OfcLock {
	_ecda := vml.NewOfcLock()
	_ecda.ExtAttr = vml.ST_ExtEdit
	_ecda.AspectratioAttr = sharedTypes.ST_TrueFalseTrue
	return _ecda
}

// Tables returns the tables defined in the footer.
func (_cbaf Footer) Tables() []Table {
	_bfbe := []Table{}
	if _cbaf._fcc == nil {
		return nil
	}
	for _, _dgdc := range _cbaf._fcc.EG_ContentBlockContent {
		for _, _fgecg := range _cbaf._aegg.tables(_dgdc) {
			_bfbe = append(_bfbe, _fgecg)
		}
	}
	return _bfbe
}

// GetColor returns the color.Color object representing the run color.
func (_feaea RunProperties) GetColor() color.Color {
	if _bggbc := _feaea._gbdb.Color; _bggbc != nil {
		_ffdae := _bggbc.ValAttr
		if _ffdae.ST_HexColorRGB != nil {
			return color.FromHex(*_ffdae.ST_HexColorRGB)
		}
	}
	return color.Color{}
}

// Strike returns true if paragraph is striked.
func (_fdbd ParagraphProperties) Strike() bool { return _cadf(_fdbd._dfaf.RPr.Strike) }

// PossibleValues returns the possible values for a FormFieldTypeDropDown.
func (_cagfc FormField) PossibleValues() []string {
	if _cagfc._cbde.DdList == nil {
		return nil
	}
	_cdcae := []string{}
	for _, _fdfc := range _cagfc._cbde.DdList.ListEntry {
		if _fdfc == nil {
			continue
		}
		_cdcae = append(_cdcae, _fdfc.ValAttr)
	}
	return _cdcae
}

// AddFootnote will create a new footnote and attach it to the Paragraph in the
// location at the end of the previous run (footnotes create their own run within
// the paragraph). The text given to the function is simply a convenience helper,
// paragraphs and runs can always be added to the text of the footnote later.
func (_dega Paragraph) AddFootnote(text string) Footnote {
	var _ffgfc int64
	if _dega._fagf.HasFootnotes() {
		for _, _cgfd := range _dega._fagf.Footnotes() {
			if _cgfd.id() > _ffgfc {
				_ffgfc = _cgfd.id()
			}
		}
		_ffgfc++
	} else {
		_ffgfc = 0
		_dega._fagf._beg = &wml.Footnotes{}
		_dega._fagf._beg.CT_Footnotes = wml.CT_Footnotes{}
		_dega._fagf._beg.Footnote = make([]*wml.CT_FtnEdn, 0)
	}
	_cdac := wml.NewCT_FtnEdn()
	_egfg := wml.NewCT_FtnEdnRef()
	_egfg.IdAttr = _ffgfc
	_dega._fagf._beg.CT_Footnotes.Footnote = append(_dega._fagf._beg.CT_Footnotes.Footnote, _cdac)
	_afcc := _dega.AddRun()
	_ccfb := _afcc.Properties()
	_ccfb.SetStyle("\u0046\u006f\u006f\u0074\u006e\u006f\u0074\u0065\u0041n\u0063\u0068\u006f\u0072")
	_afcc._adaad.EG_RunInnerContent = []*wml.EG_RunInnerContent{wml.NewEG_RunInnerContent()}
	_afcc._adaad.EG_RunInnerContent[0].FootnoteReference = _egfg
	_caga := Footnote{_dega._fagf, _cdac}
	_caga._bgcda.IdAttr = _ffgfc
	_caga._bgcda.EG_BlockLevelElts = []*wml.EG_BlockLevelElts{wml.NewEG_BlockLevelElts()}
	_cffaa := _caga.AddParagraph()
	_cffaa.Properties().SetStyle("\u0046\u006f\u006f\u0074\u006e\u006f\u0074\u0065")
	_cffaa._eagd.PPr.RPr = wml.NewCT_ParaRPr()
	_afaf := _cffaa.AddRun()
	_afaf.AddTab()
	_afaf.AddText(text)
	return _caga
}

// AddDropdownList adds dropdown list form field to the paragraph and returns it.
func (_ddf Paragraph) AddDropdownList(name string) FormField {
	_ecccf := _ddf.addFldCharsForField(name, "\u0046\u004f\u0052M\u0044\u0052\u004f\u0050\u0044\u004f\u0057\u004e")
	_ecccf._cbde.DdList = wml.NewCT_FFDDList()
	return _ecccf
}

// SizeValue returns the value of paragraph font size in points.
func (_efdag ParagraphProperties) SizeValue() float64 {
	if _fcbeg := _efdag._dfaf.RPr.Sz; _fcbeg != nil {
		_gdgf := _fcbeg.ValAttr
		if _gdgf.ST_UnsignedDecimalNumber != nil {
			return float64(*_gdgf.ST_UnsignedDecimalNumber) / 2
		}
	}
	return 0.0
}

// GetStyle returns string style of the text in watermark and format it to TextpathStyle.
func (_fefd *WatermarkText) GetStyle() vmldrawing.TextpathStyle {
	_ebeae := _fefd.getShape()
	if _fefd._bfbf != nil {
		_gefgd := _fefd._bfbf.EG_ShapeElements
		if len(_gefgd) > 0 && _gefgd[0].Textpath != nil {
			return vmldrawing.NewTextpathStyle(*_gefgd[0].Textpath.StyleAttr)
		}
	} else {
		_gbacf := _fefd.findNode(_ebeae, "\u0074\u0065\u0078\u0074\u0070\u0061\u0074\u0068")
		for _, _bfbge := range _gbacf.Attrs {
			if _bfbge.Name.Local == "\u0073\u0074\u0079l\u0065" {
				return vmldrawing.NewTextpathStyle(_bfbge.Value)
			}
		}
	}
	return vmldrawing.NewTextpathStyle("")
}

// SizeMeasure returns font with its measure which can be mm, cm, in, pt, pc or pi.
func (_bfed ParagraphProperties) SizeMeasure() string {
	if _ggcf := _bfed._dfaf.RPr.Sz; _ggcf != nil {
		_dgadd := _ggcf.ValAttr
		if _dgadd.ST_PositiveUniversalMeasure != nil {
			return *_dgadd.ST_PositiveUniversalMeasure
		}
	}
	return ""
}

// Open opens and reads a document from a file (.docx).
func Open(filename string) (*Document, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("e\u0072r\u006f\u0072\u0020\u006f\u0070\u0065\u006e\u0069n\u0067\u0020\u0025\u0073: \u0025\u0073", filename, err)
	}
	defer f.Close()
	finfo, err := os.Stat(filename)
	if err != nil {
		return nil, fmt.Errorf("e\u0072r\u006f\u0072\u0020\u006f\u0070\u0065\u006e\u0069n\u0067\u0020\u0025\u0073: \u0025\u0073", filename, err)
	}
	_ = finfo
	return Read(f, finfo.Size())
}

// SizeValue returns the value of run font size in points.
func (_gedcg RunProperties) SizeValue() float64 {
	if _begg := _gedcg._gbdb.Sz; _begg != nil {
		_caea := _begg.ValAttr
		if _caea.ST_UnsignedDecimalNumber != nil {
			return float64(*_caea.ST_UnsignedDecimalNumber) / 2
		}
	}
	return 0.0
}

// SetKeepOnOnePage controls if all lines in a paragraph are kept on the same
// page.
func (_dcbfa ParagraphStyleProperties) SetKeepOnOnePage(b bool) {
	if !b {
		_dcbfa._gfee.KeepLines = nil
	} else {
		_dcbfa._gfee.KeepLines = wml.NewCT_OnOff()
	}
}

// InsertRowAfter inserts a row after another row
func (tbl Table) InsertRowAfter(r Row) Row {
	for i, egContentRowContent := range tbl.ctTbl.EG_ContentRowContent {
		if len(egContentRowContent.Tr) > 0 && r.X() == egContentRowContent.Tr[0] {
			_egContentRowContent := wml.NewEG_ContentRowContent()
			if len(tbl.ctTbl.EG_ContentRowContent) < i+2 {
				return tbl.AddRow()
			}
			tbl.ctTbl.EG_ContentRowContent = append(tbl.ctTbl.EG_ContentRowContent, nil)
			copy(tbl.ctTbl.EG_ContentRowContent[i+2:], tbl.ctTbl.EG_ContentRowContent[i+1:])
			tbl.ctTbl.EG_ContentRowContent[i+1] = _egContentRowContent
			ctRow := wml.NewCT_Row()
			_egContentRowContent.Tr = append(_egContentRowContent.Tr, ctRow)
			return Row{tbl.doc, ctRow}
		}
	}
	return tbl.AddRow()
}

// SetLeft sets the left border to a specified type, color and thickness.
func (_aecb TableBorders) SetLeft(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_aecb._gcdf.Left = wml.NewCT_Border()
	_feadc(_aecb._gcdf.Left, t, c, thickness)
}

// SetHAlignment sets the horizontal alignment for an anchored image.
func (_ef AnchoredDrawing) SetHAlignment(h wml.WdST_AlignH) {
	_ef._dgc.PositionH.Choice = &wml.WdCT_PosHChoice{}
	_ef._dgc.PositionH.Choice.Align = h
}

// Margins allows controlling individual cell margins.
func (_ced CellProperties) Margins() CellMargins {
	if _ced._cgc.TcMar == nil {
		_ced._cgc.TcMar = wml.NewCT_TcMar()
	}
	return CellMargins{_ced._cgc.TcMar}
}

// SetColumnBandSize sets the number of Columns in the column band
func (_dadgf TableStyleProperties) SetColumnBandSize(cols int64) {
	_dadgf._degc.TblStyleColBandSize = wml.NewCT_DecimalNumber()
	_dadgf._degc.TblStyleColBandSize.ValAttr = cols
}
func _cadg(_febf string) mergeFieldInfo {
	_ggfb := []string{}
	_degbf := bytes.Buffer{}
	_geeb := -1
	for _fdda, _abbcgg := range _febf {
		switch _abbcgg {
		case ' ':
			if _degbf.Len() != 0 {
				_ggfb = append(_ggfb, _degbf.String())
			}
			_degbf.Reset()
		case '"':
			if _geeb != -1 {
				_ggfb = append(_ggfb, _febf[_geeb+1:_fdda])
				_geeb = -1
			} else {
				_geeb = _fdda
			}
		default:
			_degbf.WriteRune(_abbcgg)
		}
	}
	if _degbf.Len() != 0 {
		_ggfb = append(_ggfb, _degbf.String())
	}
	_bgcc := mergeFieldInfo{}
	for _cadab := 0; _cadab < len(_ggfb)-1; _cadab++ {
		_baad := _ggfb[_cadab]
		switch _baad {
		case "\u004d\u0045\u0052\u0047\u0045\u0046\u0049\u0045\u004c\u0044":
			_bgcc._gdfge = _ggfb[_cadab+1]
			_cadab++
		case "\u005c\u0066":
			_bgcc._cgec = _ggfb[_cadab+1]
			_cadab++
		case "\u005c\u0062":
			_bgcc._dfga = _ggfb[_cadab+1]
			_cadab++
		case "\u005c\u002a":
			switch _ggfb[_cadab+1] {
			case "\u0055\u0070\u0070e\u0072":
				_bgcc._dbfgc = true
			case "\u004c\u006f\u0077e\u0072":
				_bgcc._fgaa = true
			case "\u0043\u0061\u0070\u0073":
				_bgcc._bcdae = true
			case "\u0046\u0069\u0072\u0073\u0074\u0043\u0061\u0070":
				_bgcc._bddb = true
			}
			_cadab++
		}
	}
	return _bgcc
}

// SetText sets the watermark text.
func (_cfgbb *WatermarkText) SetText(text string) {
	_abfef := _cfgbb.getShape()
	if _cfgbb._bfbf != nil {
		_dceaa := _cfgbb._bfbf.EG_ShapeElements
		if len(_dceaa) > 0 && _dceaa[0].Textpath != nil {
			_dceaa[0].Textpath.StringAttr = &text
		}
	} else {
		_bdffc := _cfgbb.findNode(_abfef, "\u0074\u0065\u0078\u0074\u0070\u0061\u0074\u0068")
		for _fdef, _fbec := range _bdffc.Attrs {
			if _fbec.Name.Local == "\u0073\u0074\u0072\u0069\u006e\u0067" {
				_bdffc.Attrs[_fdef].Value = text
			}
		}
	}
}
func _bbgf(_fecd *wml.CT_Border, _abfce wml.ST_Border, _fbbg color.Color, _gdeed measurement.Distance) {
	_fecd.ValAttr = _abfce
	_fecd.ColorAttr = &wml.ST_HexColor{}
	if _fbbg.IsAuto() {
		_fecd.ColorAttr.ST_HexColorAuto = wml.ST_HexColorAutoAuto
	} else {
		_fecd.ColorAttr.ST_HexColorRGB = _fbbg.AsRGBString()
	}
	if _gdeed != measurement.Zero {
		_fecd.SzAttr = unioffice.Uint64(uint64(_gdeed / measurement.Point * 8))
	}
}

// AddDrawingInline adds an inline drawing from an ImageRef.
func (_bcff Run) AddDrawingInline(img common.ImageRef) (InlineDrawing, error) {
	_dbcab := _bcff.newIC()
	_dbcab.Drawing = wml.NewCT_Drawing()
	_bbda := wml.NewWdInline()
	_ffdaa := InlineDrawing{_bcff._dbddf, _bbda}
	_bbda.CNvGraphicFramePr = dml.NewCT_NonVisualGraphicFrameProperties()
	_dbcab.Drawing.Inline = append(_dbcab.Drawing.Inline, _bbda)
	_bbda.Graphic = dml.NewGraphic()
	_bbda.Graphic.GraphicData = dml.NewCT_GraphicalObjectData()
	_bbda.Graphic.GraphicData.UriAttr = "\u0068\u0074\u0074\u0070\u003a\u002f/\u0073\u0063\u0068e\u006d\u0061\u0073.\u006f\u0070\u0065\u006e\u0078\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002e\u006frg\u002f\u0064\u0072\u0061\u0077\u0069\u006e\u0067\u006d\u006c\u002f\u0032\u0030\u0030\u0036\u002f\u0070\u0069\u0063\u0074\u0075\u0072\u0065"
	_bbda.DistTAttr = unioffice.Uint32(0)
	_bbda.DistLAttr = unioffice.Uint32(0)
	_bbda.DistBAttr = unioffice.Uint32(0)
	_bbda.DistRAttr = unioffice.Uint32(0)
	_bbda.Extent.CxAttr = int64(float64(img.Size().X*measurement.Pixel72) / measurement.EMU)
	_bbda.Extent.CyAttr = int64(float64(img.Size().Y*measurement.Pixel72) / measurement.EMU)
	_bagde := 0x7FFFFFFF & rand.Uint32()
	_bbda.DocPr.IdAttr = _bagde
	_cegab := picture.NewPic()
	_cegab.NvPicPr.CNvPr.IdAttr = _bagde
	_ffgeb := img.RelID()
	if _ffgeb == "" {
		return _ffdaa, errors.New("\u0063\u006f\u0075\u006c\u0064\u006e\u0027\u0074\u0020\u0066\u0069\u006e\u0064\u0020\u0072\u0065\u0066\u0065\u0072\u0065n\u0063\u0065\u0020\u0074\u006f\u0020\u0069\u006d\u0061g\u0065\u0020\u0077\u0069\u0074\u0068\u0069\u006e\u0020\u0064\u006f\u0063\u0075m\u0065\u006e\u0074\u0020\u0072\u0065l\u0061\u0074\u0069o\u006e\u0073")
	}
	_bbda.Graphic.GraphicData.Any = append(_bbda.Graphic.GraphicData.Any, _cegab)
	_cegab.BlipFill = dml.NewCT_BlipFillProperties()
	_cegab.BlipFill.Blip = dml.NewCT_Blip()
	_cegab.BlipFill.Blip.EmbedAttr = &_ffgeb
	_cegab.BlipFill.Stretch = dml.NewCT_StretchInfoProperties()
	_cegab.BlipFill.Stretch.FillRect = dml.NewCT_RelativeRect()
	_cegab.SpPr = dml.NewCT_ShapeProperties()
	_cegab.SpPr.Xfrm = dml.NewCT_Transform2D()
	_cegab.SpPr.Xfrm.Off = dml.NewCT_Point2D()
	_cegab.SpPr.Xfrm.Off.XAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_cegab.SpPr.Xfrm.Off.YAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	_cegab.SpPr.Xfrm.Ext = dml.NewCT_PositiveSize2D()
	_cegab.SpPr.Xfrm.Ext.CxAttr = int64(img.Size().X * measurement.Point)
	_cegab.SpPr.Xfrm.Ext.CyAttr = int64(img.Size().Y * measurement.Point)
	_cegab.SpPr.PrstGeom = dml.NewCT_PresetGeometry2D()
	_cegab.SpPr.PrstGeom.PrstAttr = dml.ST_ShapeTypeRect
	return _ffdaa, nil
}

// StructuredDocumentTags returns the structured document tags in the document
// which are commonly used in document templates.
func (_fcfb *Document) StructuredDocumentTags() []StructuredDocumentTag {
	_adacd := []StructuredDocumentTag{}
	for _, _fgg := range _fcfb.doc.Body.EG_BlockLevelElts {
		for _, _gfc := range _fgg.EG_ContentBlockContent {
			if _gfc.Sdt != nil {
				_adacd = append(_adacd, StructuredDocumentTag{_fcfb, _gfc.Sdt})
			}
		}
	}
	return _adacd
}

// OpenTemplate opens a document, removing all content so it can be used as a
// template.  Since Word removes unused styles from a document upon save, to
// create a template in Word add a paragraph with every style of interest.  When
// opened with OpenTemplate the document's styles will be available but the
// content will be gone.
func OpenTemplate(filename string) (*Document, error) {
	doc, err := Open(filename)
	if err != nil {
		return nil, err
	}
	doc.doc.Body = wml.NewCT_Body()
	return doc, nil
}

// AddTabStop adds a tab stop to the paragraph.
func (_dbcg ParagraphStyleProperties) AddTabStop(position measurement.Distance, justificaton wml.ST_TabJc, leader wml.ST_TabTlc) {
	if _dbcg._gfee.Tabs == nil {
		_dbcg._gfee.Tabs = wml.NewCT_Tabs()
	}
	_dege := wml.NewCT_TabStop()
	_dege.LeaderAttr = leader
	_dege.ValAttr = justificaton
	_dege.PosAttr.Int64 = unioffice.Int64(int64(position / measurement.Twips))
	_dbcg._gfee.Tabs.Tab = append(_dbcg._gfee.Tabs.Tab, _dege)
}

// X returns the inner wrapped XML type.
func (tbl Table) X() *wml.CT_Tbl {
	return tbl.ctTbl
}

// GetChartSpaceByRelId returns a *crt.ChartSpace with the associated relation ID in the
// document.
func (_cbad *Document) GetChartSpaceByRelId(relId string) *dmlChart.ChartSpace {
	_cbgd := _cbad._dab.GetTargetByRelId(relId)
	for _, _gdffa := range _cbad._caf {
		if _cbgd == _gdffa.Target() {
			return _gdffa._ffb
		}
	}
	return nil
}

// AddRow adds a row to a table.
func (tbl Table) AddRow() Row {
	egContentRowContent := wml.NewEG_ContentRowContent()
	tbl.ctTbl.EG_ContentRowContent = append(tbl.ctTbl.EG_ContentRowContent, egContentRowContent)
	ctRow := wml.NewCT_Row()
	egContentRowContent.Tr = append(egContentRowContent.Tr, ctRow)
	return Row{tbl.doc, ctRow}
}

func (_fggac *WatermarkText) findNode(_gcadd *unioffice.XSDAny, _ddgf string) *unioffice.XSDAny {
	for _, _fdedf := range _gcadd.Nodes {
		if _fdedf.XMLName.Local == _ddgf {
			return _fdedf
		}
	}
	return nil
}

// SetStartPct sets the cell start margin
func (_beae CellMargins) SetStartPct(pct float64) {
	_beae._cdae.Start = wml.NewCT_TblWidth()
	_aff(_beae._cdae.Start, pct)
}

// GetWrapPathLineTo return wrapPath lineTo value.
func (_efb AnchorDrawWrapOptions) GetWrapPathLineTo() []*dml.CT_Point2D { return _efb._cbf }
func (_ggecb *WatermarkPicture) getShapeImagedata() *unioffice.XSDAny {
	return _ggecb.getInnerElement("\u0069m\u0061\u0067\u0065\u0064\u0061\u0074a")
}

// SetInsideVertical sets the interior vertical borders to a specified type, color and thickness.
func (_acbfb TableBorders) SetInsideVertical(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_acbfb._gcdf.InsideV = wml.NewCT_Border()
	_feadc(_acbfb._gcdf.InsideV, t, c, thickness)
}

// SetStartIndent controls the start indentation.
func (_cfcf ParagraphProperties) SetStartIndent(m measurement.Distance) {
	if _cfcf._dfaf.Ind == nil {
		_cfcf._dfaf.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_cfcf._dfaf.Ind.StartAttr = nil
	} else {
		_cfcf._dfaf.Ind.StartAttr = &wml.ST_SignedTwipsMeasure{}
		_cfcf._dfaf.Ind.StartAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// GetDocRelTargetByID returns TargetAttr of document relationship given its IdAttr.
func (_bdc *Document) GetDocRelTargetByID(idAttr string) string {
	for _, _ddca := range _bdc._dab.X().Relationship {
		if _ddca.IdAttr == idAttr {
			return _ddca.TargetAttr
		}
	}
	return ""
}

// ParagraphStyleProperties is the styling information for a paragraph.
type ParagraphStyleProperties struct{ _gfee *wml.CT_PPrGeneral }

// Clear removes all of the content from within a run.
func (_cgbg Run) Clear() { _cgbg._adaad.EG_RunInnerContent = nil }

// AddParagraph adds a paragraph to the header.
func (_fgbe Header) AddParagraph() Paragraph {
	_gbeaa := wml.NewEG_ContentBlockContent()
	_fgbe._deae.EG_ContentBlockContent = append(_fgbe._deae.EG_ContentBlockContent, _gbeaa)
	_fbcd := wml.NewCT_P()
	_gbeaa.P = append(_gbeaa.P, _fbcd)
	return Paragraph{_fgbe._dbagd, _fbcd}
}

// SetKeepWithNext controls if this paragraph should be kept with the next.
func (_afgcc ParagraphProperties) SetKeepWithNext(b bool) {
	if !b {
		_afgcc._dfaf.KeepNext = nil
	} else {
		_afgcc._dfaf.KeepNext = wml.NewCT_OnOff()
	}
}

// Clear resets the numbering.
func (_ddacg Numbering) Clear() {
	_ddacg._cbag.AbstractNum = nil
	_ddacg._cbag.Num = nil
	_ddacg._cbag.NumIdMacAtCleanup = nil
	_ddacg._cbag.NumPicBullet = nil
}

// SetColor sets the text color.
func (_ffce RunProperties) SetColor(c color.Color) {
	_ffce._gbdb.Color = wml.NewCT_Color()
	_ffce._gbdb.Color.ValAttr.ST_HexColorRGB = c.AsRGBString()
}

// SetValue sets the value of a FormFieldTypeText or FormFieldTypeDropDown. For
// FormFieldTypeDropDown, the value must be one of the fields possible values.
func (_bccb FormField) SetValue(v string) {
	if _bccb._cbde.DdList != nil {
		for _fbfgc, _agdd := range _bccb.PossibleValues() {
			if _agdd == v {
				_bccb._cbde.DdList.Result = wml.NewCT_DecimalNumber()
				_bccb._cbde.DdList.Result.ValAttr = int64(_fbfgc)
				break
			}
		}
	} else if _bccb._cbde.TextInput != nil {
		_bccb._gcbd.T = wml.NewCT_Text()
		_bccb._gcbd.T.Content = v
	}
}
func (_cefc Styles) initializeDocDefaults() {
	_cefc._abca.DocDefaults = wml.NewCT_DocDefaults()
	_cefc._abca.DocDefaults.RPrDefault = wml.NewCT_RPrDefault()
	_cefc._abca.DocDefaults.RPrDefault.RPr = wml.NewCT_RPr()
	_gaefc := RunProperties{_cefc._abca.DocDefaults.RPrDefault.RPr}
	_gaefc.SetSize(12 * measurement.Point)
	_gaefc.Fonts().SetASCIITheme(wml.ST_ThemeMajorAscii)
	_gaefc.Fonts().SetEastAsiaTheme(wml.ST_ThemeMajorEastAsia)
	_gaefc.Fonts().SetHANSITheme(wml.ST_ThemeMajorHAnsi)
	_gaefc.Fonts().SetCSTheme(wml.ST_ThemeMajorBidi)
	_gaefc.X().Lang = wml.NewCT_Language()
	_gaefc.X().Lang.ValAttr = unioffice.String("\u0065\u006e\u002dU\u0053")
	_gaefc.X().Lang.EastAsiaAttr = unioffice.String("\u0065\u006e\u002dU\u0053")
	_gaefc.X().Lang.BidiAttr = unioffice.String("\u0061\u0072\u002dS\u0041")
	_cefc._abca.DocDefaults.PPrDefault = wml.NewCT_PPrDefault()
}

// SetText sets the text to be used in bullet mode.
func (_gcaab NumberingLevel) SetText(t string) {
	if t == "" {
		_gcaab.lvl.LvlText = nil
	} else {
		_gcaab.lvl.LvlText = wml.NewCT_LevelText()
		_gcaab.lvl.LvlText.ValAttr = unioffice.String(t)
	}
}

// SetBetween sets the between border to a specified type, color and thickness between paragraph.
func (_cbbb ParagraphBorders) SetBetween(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_cbbb._fdge.Between = wml.NewCT_Border()
	_bbgf(_cbbb._fdge.Between, t, c, thickness)
}

// Table is a table within a document.
type Table struct {
	doc *Document
	ctTbl *wml.CT_Tbl
}

func (_ceee *WatermarkPicture) findNode(_agfab *unioffice.XSDAny, _bgge string) *unioffice.XSDAny {
	for _, _dcecd := range _agfab.Nodes {
		if _dcecd.XMLName.Local == _bgge {
			return _dcecd
		}
	}
	return nil
}

// X returns the inner wrapped XML type.
func (_daffe TableLook) X() *wml.CT_TblLook { return _daffe.ctTblLook }

// Shadow returns true if paragraph shadow is on.
func (_afgad ParagraphProperties) Shadow() bool { return _cadf(_afgad._dfaf.RPr.Shadow) }

// SetLinkedStyle sets the style that this style is linked to.
func (_ccbgg Style) SetLinkedStyle(name string) {
	if name == "" {
		_ccbgg._gaege.Link = nil
	} else {
		_ccbgg._gaege.Link = wml.NewCT_String()
		_ccbgg._gaege.Link.ValAttr = name
	}
}

// SetSize sets the size of the displayed image on the page.
func (_gdbec InlineDrawing) SetSize(w, h measurement.Distance) {
	_gdbec._ecag.Extent.CxAttr = int64(float64(w*measurement.Pixel72) / measurement.EMU)
	_gdbec._ecag.Extent.CyAttr = int64(float64(h*measurement.Pixel72) / measurement.EMU)
}

// SetHighlight highlights text in a specified color.
func (_egec RunProperties) SetHighlight(c wml.ST_HighlightColor) {
	_egec._gbdb.Highlight = wml.NewCT_Highlight()
	_egec._gbdb.Highlight.ValAttr = c
}

// NewSettings constructs a new empty Settings
func NewSettings() Settings {
	_afba := wml.NewSettings()
	_afba.Compat = wml.NewCT_Compat()
	_fcfa := wml.NewCT_CompatSetting()
	_fcfa.NameAttr = unioffice.String("\u0063\u006f\u006d\u0070\u0061\u0074\u0069\u0062\u0069\u006c\u0069\u0074y\u004d\u006f\u0064\u0065")
	_fcfa.UriAttr = unioffice.String("h\u0074\u0074\u0070\u003a\u002f\u002f\u0073\u0063\u0068\u0065\u006d\u0061\u0073\u002e\u006d\u0069\u0063\u0072o\u0073\u006f\u0066\u0074\u002e\u0063\u006f\u006d\u002f\u006fff\u0069\u0063\u0065/\u0077o\u0072\u0064")
	_fcfa.ValAttr = unioffice.String("\u0031\u0035")
	_afba.Compat.CompatSetting = append(_afba.Compat.CompatSetting, _fcfa)
	return Settings{_afba}
}

// AddRun adds a run of text to a hyperlink. This is the text that will be linked.
func (_badb HyperLink) AddRun() Run {
	_gacf := wml.NewEG_ContentRunContent()
	_badb._baaf.EG_ContentRunContent = append(_badb._baaf.EG_ContentRunContent, _gacf)
	_ggbaa := wml.NewCT_R()
	_gacf.R = _ggbaa
	return Run{_badb._acbfa, _ggbaa}
}

// SetAll sets all of the borders to a given value.
func (_ged CellBorders) SetAll(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_ged.SetBottom(t, c, thickness)
	_ged.SetLeft(t, c, thickness)
	_ged.SetRight(t, c, thickness)
	_ged.SetTop(t, c, thickness)
	_ged.SetInsideHorizontal(t, c, thickness)
	_ged.SetInsideVertical(t, c, thickness)
}

// StyleID returns the style ID.
func (_cfaf Style) StyleID() string {
	if _cfaf._gaege.StyleIdAttr == nil {
		return ""
	}
	return *_cfaf._gaege.StyleIdAttr
}

// Footers returns the footers defined in the document.
func (_eggd *Document) Footers() []Footer {
	_gfa := []Footer{}
	for _, _afad := range _eggd._aba {
		_gfa = append(_gfa, Footer{_eggd, _afad})
	}
	return _gfa
}

// SetVerticalBanding controls the conditional formatting for vertical banding.
func (_bbecb TableLook) SetVerticalBanding(on bool) {
	if !on {
		_bbecb.ctTblLook.NoVBandAttr = &sharedTypes.ST_OnOff{}
		_bbecb.ctTblLook.NoVBandAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	} else {
		_bbecb.ctTblLook.NoVBandAttr = &sharedTypes.ST_OnOff{}
		_bbecb.ctTblLook.NoVBandAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	}
}

// X returns the inner wrapped XML type.
func (_egcb Header) X() *wml.Hdr { return _egcb._deae }

// SetTextWrapBehindText sets the text wrap to behind text.
func (_eg AnchoredDrawing) SetTextWrapBehindText() {
	_eg._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_eg._dgc.Choice.WrapNone = wml.NewWdCT_WrapNone()
	_eg._dgc.BehindDocAttr = true
	_eg._dgc.LayoutInCellAttr = true
	_eg._dgc.AllowOverlapAttr = true
}

// SetBeforeSpacing sets spacing above paragraph.
func (_cgafc Paragraph) SetBeforeSpacing(d measurement.Distance) {
	_cgafc.ensurePPr()
	if _cgafc._eagd.PPr.Spacing == nil {
		_cgafc._eagd.PPr.Spacing = wml.NewCT_Spacing()
	}
	_dcbaf := _cgafc._eagd.PPr.Spacing
	_dcbaf.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
	_dcbaf.BeforeAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(d / measurement.Twips))
}

// Control returns an *axcontrol.Control object contained in the run or the nil value in case of no controls.
func (_gacc Run) Control() *axcontrol.Control {
	if _deaeb := _gacc._adaad.EG_RunInnerContent; _deaeb != nil {
		if _baadd := _deaeb[0].Object; _baadd != nil {
			if _degg := _baadd.Choice; _degg != nil {
				if _dfcfe := _degg.Control; _dfcfe != nil {
					if _dfcfe.IdAttr != nil {
						_eede := _gacc._dbddf.GetDocRelTargetByID(*_dfcfe.IdAttr)
						for _, _cgga := range _gacc._dbddf._gga {
							if _eede == _cgga.TargetAttr {
								return _cgga
							}
						}
					}
				}
			}
		}
	}
	return nil
}

// SetVerticalAlignment sets the vertical alignment of content within a table cell.
func (_def CellProperties) SetVerticalAlignment(align wml.ST_VerticalJc) {
	if align == wml.ST_VerticalJcUnset {
		_def._cgc.VAlign = nil
	} else {
		_def._cgc.VAlign = wml.NewCT_VerticalJc()
		_def._cgc.VAlign.ValAttr = align
	}
}

// Type returns the type of the style.
func (_bbfae Style) Type() wml.ST_StyleType { return _bbfae._gaege.TypeAttr }
func (_dcd *Document) validateBookmarks() error {
	_dbdb := make(map[string]struct{})
	for _, _bggb := range _dcd.Bookmarks() {
		if _, _agb := _dbdb[_bggb.Name()]; _agb {
			return fmt.Errorf("d\u0075\u0070\u006c\u0069\u0063\u0061t\u0065\u0020\u0062\u006f\u006f\u006b\u006d\u0061\u0072k\u0020\u0025\u0073 \u0066o\u0075\u006e\u0064", _bggb.Name())
		}
		_dbdb[_bggb.Name()] = struct{}{}
	}
	return nil
}

// SetWindowControl controls if the first or last line of the paragraph is
// allowed to dispay on a separate page.
func (_dbcf ParagraphProperties) SetWindowControl(b bool) {
	if !b {
		_dbcf._dfaf.WidowControl = nil
	} else {
		_dbcf._dfaf.WidowControl = wml.NewCT_OnOff()
	}
}
func _cgcae() *vml.Formulas {
	_gacd := vml.NewFormulas()
	_gacd.F = []*vml.CT_F{vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0023\u0030\u0020\u0030\u00201\u0030\u0038\u0030\u0030"), vmldrawing.CreateFormula("p\u0072\u006f\u0064\u0020\u0023\u0030\u0020\u0032\u0020\u0031"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0032\u0031\u0036\u0030\u0030 \u0030\u0020\u0040\u0031"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0030\u0020\u0030\u0020\u0040\u0032"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0032\u0031\u0036\u0030\u0030 \u0030\u0020\u0040\u0033"), vmldrawing.CreateFormula("\u0069\u0066\u0020\u0040\u0030\u0020\u0040\u0033\u0020\u0030"), vmldrawing.CreateFormula("\u0069\u0066\u0020\u0040\u0030\u0020\u0032\u0031\u00360\u0030\u0020\u0040\u0031"), vmldrawing.CreateFormula("\u0069\u0066\u0020\u0040\u0030\u0020\u0030\u0020\u0040\u0032"), vmldrawing.CreateFormula("\u0069\u0066\u0020\u0040\u0030\u0020\u0040\u0034\u00202\u0031\u0036\u0030\u0030"), vmldrawing.CreateFormula("\u006di\u0064\u0020\u0040\u0035\u0020\u00406"), vmldrawing.CreateFormula("\u006di\u0064\u0020\u0040\u0038\u0020\u00405"), vmldrawing.CreateFormula("\u006di\u0064\u0020\u0040\u0037\u0020\u00408"), vmldrawing.CreateFormula("\u006di\u0064\u0020\u0040\u0036\u0020\u00407"), vmldrawing.CreateFormula("s\u0075\u006d\u0020\u0040\u0036\u0020\u0030\u0020\u0040\u0035")}
	return _gacd
}

// Pict returns the pict object.
func (_ebaae *WatermarkPicture) Pict() *wml.CT_Picture { return _ebaae._cdff }

// SetWrapPathLineTo sets wrapPath lineTo value.
func (_aaa AnchorDrawWrapOptions) SetWrapPathLineTo(coordinates []*dml.CT_Point2D) {
	_aaa._cbf = coordinates
}

// Footnote is an individual footnote reference within the document.
type Footnote struct {
	_gffg  *Document
	_bgcda *wml.CT_FtnEdn
}

// ParagraphStyles returns only the paragraph styles.
func (_gaegc Styles) ParagraphStyles() []Style {
	_bfgcg := []Style{}
	for _, _dabb := range _gaegc._abca.Style {
		if _dabb.TypeAttr != wml.ST_StyleTypeParagraph {
			continue
		}
		_bfgcg = append(_bfgcg, Style{_dabb})
	}
	return _bfgcg
}

// SetValue sets the width value.
func (_gedgf TableWidth) SetValue(m measurement.Distance) {
	_gedgf._egbb.WAttr = &wml.ST_MeasurementOrPercent{}
	_gedgf._egbb.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_gedgf._egbb.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(m / measurement.Twips))
	_gedgf._egbb.TypeAttr = wml.ST_TblWidthDxa
}

// SetBasedOn sets the style that this style is based on.
func (_dddf Style) SetBasedOn(name string) {
	if name == "" {
		_dddf._gaege.BasedOn = nil
	} else {
		_dddf._gaege.BasedOn = wml.NewCT_String()
		_dddf._gaege.BasedOn.ValAttr = name
	}
}

// SetTextWrapInFrontOfText sets the text wrap to in front of text.
func (_eec AnchoredDrawing) SetTextWrapInFrontOfText() {
	_eec._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_eec._dgc.Choice.WrapNone = wml.NewWdCT_WrapNone()
	_eec._dgc.BehindDocAttr = false
	_eec._dgc.LayoutInCellAttr = true
	_eec._dgc.AllowOverlapAttr = true
}

// SetLastColumn controls the conditional formatting for the last column in a table.
func (_beda TableLook) SetLastColumn(on bool) {
	if !on {
		_beda.ctTblLook.LastColumnAttr = &sharedTypes.ST_OnOff{}
		_beda.ctTblLook.LastColumnAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	} else {
		_beda.ctTblLook.LastColumnAttr = &sharedTypes.ST_OnOff{}
		_beda.ctTblLook.LastColumnAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	}
}
func (_gagc *Document) onNewRelationship(_gfgce *zippkg.DecodeMap, _gfff, _ffe string, _gec []*zip.File, _ccg *relationships.Relationship, _fgb zippkg.Target) error {
	_ecgbc := unioffice.DocTypeDocument
	switch _ffe {
	case unioffice.OfficeDocumentType, unioffice.OfficeDocumentTypeStrict:
		_gagc.doc = wml.NewDocument()
		_gfgce.AddTarget(_gfff, _gagc.doc, _ffe, 0)
		_gfgce.AddTarget(zippkg.RelationsPathFor(_gfff), _gagc._dab.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.CorePropertiesType:
		_gfgce.AddTarget(_gfff, _gagc.CoreProperties.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.CustomPropertiesType:
		_gfgce.AddTarget(_gfff, _gagc.CustomProperties.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.ExtendedPropertiesType, unioffice.ExtendedPropertiesTypeStrict:
		_gfgce.AddTarget(_gfff, _gagc.AppProperties.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.ThumbnailType, unioffice.ThumbnailTypeStrict:
		for _dgb, _bccf := range _gec {
			if _bccf == nil {
				continue
			}
			if _bccf.Name == _gfff {
				_cgf, _efgeg := _bccf.Open()
				if _efgeg != nil {
					return fmt.Errorf("e\u0072\u0072\u006f\u0072\u0020\u0072e\u0061\u0064\u0069\u006e\u0067\u0020\u0074\u0068\u0075m\u0062\u006e\u0061i\u006c:\u0020\u0025\u0073", _efgeg)
				}
				_gagc.Thumbnail, _, _efgeg = image.Decode(_cgf)
				_cgf.Close()
				if _efgeg != nil {
					return fmt.Errorf("\u0065\u0072\u0072\u006fr\u0020\u0064\u0065\u0063\u006f\u0064\u0069\u006e\u0067\u0020t\u0068u\u006d\u0062\u006e\u0061\u0069\u006c\u003a \u0025\u0073", _efgeg)
				}
				_gec[_dgb] = nil
			}
		}
	case unioffice.SettingsType, unioffice.SettingsTypeStrict:
		_gfgce.AddTarget(_gfff, _gagc.Settings.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.NumberingType, unioffice.NumberingTypeStrict:
		_gagc.Numbering = NewNumbering()
		_gfgce.AddTarget(_gfff, _gagc.Numbering.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.StylesType, unioffice.StylesTypeStrict:
		_gagc.Styles.Clear()
		_gfgce.AddTarget(_gfff, _gagc.Styles.X(), _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.HeaderType, unioffice.HeaderTypeStrict:
		_cacc := wml.NewHdr()
		_gfgce.AddTarget(_gfff, _cacc, _ffe, uint32(len(_gagc._geb)))
		_gagc._geb = append(_gagc._geb, _cacc)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc._geb))
		_dbag := common.NewRelationships()
		_gfgce.AddTarget(zippkg.RelationsPathFor(_gfff), _dbag.X(), _ffe, 0)
		_gagc._cbfd = append(_gagc._cbfd, _dbag)
	case unioffice.FooterType, unioffice.FooterTypeStrict:
		_feadg := wml.NewFtr()
		_gfgce.AddTarget(_gfff, _feadg, _ffe, uint32(len(_gagc._aba)))
		_gagc._aba = append(_gagc._aba, _feadg)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc._aba))
		_gabf := common.NewRelationships()
		_gfgce.AddTarget(zippkg.RelationsPathFor(_gfff), _gabf.X(), _ffe, 0)
		_gagc._fdf = append(_gagc._fdf, _gabf)
	case unioffice.ThemeType, unioffice.ThemeTypeStrict:
		_fac := dml.NewTheme()
		_gfgce.AddTarget(_gfff, _fac, _ffe, uint32(len(_gagc._ffbc)))
		_gagc._ffbc = append(_gagc._ffbc, _fac)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc._ffbc))
	case unioffice.WebSettingsType, unioffice.WebSettingsTypeStrict:
		_gagc._gbe = wml.NewWebSettings()
		_gfgce.AddTarget(_gfff, _gagc._gbe, _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.FontTableType, unioffice.FontTableTypeStrict:
		_gagc._eaa = wml.NewFonts()
		_gfgce.AddTarget(_gfff, _gagc._eaa, _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.EndNotesType, unioffice.EndNotesTypeStrict:
		_gagc._ccb = wml.NewEndnotes()
		_gfgce.AddTarget(_gfff, _gagc._ccb, _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.FootNotesType, unioffice.FootNotesTypeStrict:
		_gagc._beg = wml.NewFootnotes()
		_gfgce.AddTarget(_gfff, _gagc._beg, _ffe, 0)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, 0)
	case unioffice.ImageType, unioffice.ImageTypeStrict:
		var _fgdf common.ImageRef
		for _fbb, _cdca := range _gec {
			if _cdca == nil {
				continue
			}
			if _cdca.Name == _gfff {
				_fcea, _edfg := zippkg.ExtractToDiskTmp(_cdca, _gagc.TmpPath)
				if _edfg != nil {
					return _edfg
				}
				_eebd, _edfg := common.ImageFromStorage(_fcea)
				if _edfg != nil {
					return _edfg
				}
				_fgdf = common.MakeImageRef(_eebd, &_gagc.DocBase, _gagc._dab)
				_gec[_fbb] = nil
			}
		}
		if _fgdf.Format() != "" {
			_fdce := "\u002e" + strings.ToLower(_fgdf.Format())
			_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc.Images)+1)
			if _gfcg := filepath.Ext(_ccg.TargetAttr); _gfcg != _fdce {
				_ccg.TargetAttr = _ccg.TargetAttr[0:len(_ccg.TargetAttr)-len(_gfcg)] + _fdce
			}
			_fgdf.SetTarget("\u0077\u006f\u0072d\u002f" + _ccg.TargetAttr)
			_gagc.Images = append(_gagc.Images, _fgdf)
		}
	case unioffice.ControlType, unioffice.ControlTypeStrict:
		_beafe := activeX.NewOcx()
		_gfbf := unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc._gga)+1)
		_bcb := "\u0077\u006f\u0072d\u002f" + _gfbf[:len(_gfbf)-4] + "\u002e\u0062\u0069\u006e"
		for _gagb, _bdd := range _gec {
			if _bdd == nil {
				continue
			}
			if _bdd.Name == _bcb {
				_fecg, _dbae := zippkg.ExtractToDiskTmp(_bdd, _gagc.TmpPath)
				if _dbae != nil {
					return _dbae
				}
				_acc, _dbae := axcontrol.ImportFromFile(_fecg)
				if _dbae == nil {
					_acc.TargetAttr = _gfbf
					_acc.Ocx = _beafe
					_gagc._gga = append(_gagc._gga, _acc)
					_gfgce.AddTarget(_gfff, _beafe, _ffe, uint32(len(_gagc._gga)))
					_ccg.TargetAttr = _gfbf
					_gec[_gagb] = nil
				} else {
					logger.Log.Debug("c\u0061\u006e\u006e\u006f\u0074\u0020r\u0065\u0061\u0064\u0020\u0062\u0069\u006e\u0020\u0066i\u006c\u0065\u003a \u0025s\u0020\u0025\u0073", _bcb, _dbae.Error())
				}
				break
			}
		}
	case unioffice.ChartType:
		_cbba := chart{_ffb: dmlChart.NewChartSpace()}
		_aagb := uint32(len(_gagc._caf))
		_gfgce.AddTarget(_gfff, _cbba._ffb, _ffe, _aagb)
		_gagc._caf = append(_gagc._caf, &_cbba)
		_ccg.TargetAttr = unioffice.RelativeFilename(_ecgbc, _fgb.Typ, _ffe, len(_gagc._caf))
		_cbba._cce = _ccg.TargetAttr
	default:
		logger.Log.Debug("\u0075\u006e\u0073\u0075\u0070p\u006f\u0072\u0074\u0065\u0064\u0020\u0072\u0065\u006c\u0061\u0074\u0069\u006fn\u0073\u0068\u0069\u0070\u0020\u0074\u0079\u0070\u0065\u003a\u0020\u0025\u0073\u0020\u0074\u0067\u0074\u003a\u0020\u0025\u0073", _ffe, _gfff)
	}
	return nil
}
func (_bbgb *Document) insertParagraph(_efca Paragraph, _bac bool) Paragraph {
	if _bbgb.doc.Body == nil {
		return _bbgb.AddParagraph()
	}
	_beca := _efca.X()
	for _, _eccca := range _bbgb.doc.Body.EG_BlockLevelElts {
		for _, _abgc := range _eccca.EG_ContentBlockContent {
			for _eecc, _eeba := range _abgc.P {
				if _eeba == _beca {
					_fgdc := wml.NewCT_P()
					_abgc.P = append(_abgc.P, nil)
					if _bac {
						copy(_abgc.P[_eecc+1:], _abgc.P[_eecc:])
						_abgc.P[_eecc] = _fgdc
					} else {
						copy(_abgc.P[_eecc+2:], _abgc.P[_eecc+1:])
						_abgc.P[_eecc+1] = _fgdc
					}
					return Paragraph{_bbgb, _fgdc}
				}
			}
			for _, _gdbe := range _abgc.Tbl {
				for _, _bdb := range _gdbe.EG_ContentRowContent {
					for _, _bfbb := range _bdb.Tr {
						for _, _afga := range _bfbb.EG_ContentCellContent {
							for _, _gbgb := range _afga.Tc {
								for _, _eeed := range _gbgb.EG_BlockLevelElts {
									for _, _dfbgf := range _eeed.EG_ContentBlockContent {
										for _aggfe, _adce := range _dfbgf.P {
											if _adce == _beca {
												_acda := wml.NewCT_P()
												_dfbgf.P = append(_dfbgf.P, nil)
												if _bac {
													copy(_dfbgf.P[_aggfe+1:], _dfbgf.P[_aggfe:])
													_dfbgf.P[_aggfe] = _acda
												} else {
													copy(_dfbgf.P[_aggfe+2:], _dfbgf.P[_aggfe+1:])
													_dfbgf.P[_aggfe+1] = _acda
												}
												return Paragraph{_bbgb, _acda}
											}
										}
									}
								}
							}
						}
					}
				}
			}
			if _abgc.Sdt != nil && _abgc.Sdt.SdtContent != nil && _abgc.Sdt.SdtContent.P != nil {
				for _efcf, _acb := range _abgc.Sdt.SdtContent.P {
					if _acb == _beca {
						_afgf := wml.NewCT_P()
						_abgc.Sdt.SdtContent.P = append(_abgc.Sdt.SdtContent.P, nil)
						if _bac {
							copy(_abgc.Sdt.SdtContent.P[_efcf+1:], _abgc.Sdt.SdtContent.P[_efcf:])
							_abgc.Sdt.SdtContent.P[_efcf] = _afgf
						} else {
							copy(_abgc.Sdt.SdtContent.P[_efcf+2:], _abgc.Sdt.SdtContent.P[_efcf+1:])
							_abgc.Sdt.SdtContent.P[_efcf+1] = _afgf
						}
						return Paragraph{_bbgb, _afgf}
					}
				}
			}
		}
	}
	return _bbgb.AddParagraph()
}

// SetWidthAuto sets the the table width to automatic.
func (_ggff TableProperties) SetWidthAuto() {
	_ggff._efag.TblW = wml.NewCT_TblWidth()
	_ggff._efag.TblW.TypeAttr = wml.ST_TblWidthAuto
}

// SetTarget sets the URL target of the hyperlink.
func (_ceaag HyperLink) SetTarget(url string) {
	_cdeb := _ceaag._acbfa.AddHyperlink(url)
	_ceaag._baaf.IdAttr = unioffice.String(common.Relationship(_cdeb).ID())
	_ceaag._baaf.AnchorAttr = nil
}

// AddImage adds an image to the document package, returning a reference that
// can be used to add the image to a run and place it in the document contents.
func (_gbadf *Document) AddImage(i common.Image) (common.ImageRef, error) {
	_cea := common.MakeImageRef(i, &_gbadf.DocBase, _gbadf._dab)
	if i.Data == nil && i.Path == "" {
		return _cea, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0064\u0061t\u0061\u0020\u006f\u0072\u0020\u0061\u0020\u0070\u0061\u0074\u0068")
	}
	if i.Format == "" {
		return _cea, errors.New("\u0069\u006d\u0061\u0067\u0065\u0020\u006d\u0075\u0073\u0074 \u0068\u0061\u0076\u0065\u0020\u0061\u0020v\u0061\u006c\u0069\u0064\u0020\u0066\u006f\u0072\u006d\u0061\u0074")
	}
	if i.Size.X == 0 || i.Size.Y == 0 {
		return _cea, errors.New("\u0069\u006d\u0061\u0067e\u0020\u006d\u0075\u0073\u0074\u0020\u0068\u0061\u0076\u0065 \u0061 \u0076\u0061\u006c\u0069\u0064\u0020\u0073i\u007a\u0065")
	}
	if i.Path != "" {
		_aceg := tempstorage.Add(i.Path)
		if _aceg != nil {
			return _cea, _aceg
		}
	}
	_gbadf.Images = append(_gbadf.Images, _cea)
	_fbed := fmt.Sprintf("\u006d\u0065d\u0069\u0061\u002fi\u006d\u0061\u0067\u0065\u0025\u0064\u002e\u0025\u0073", len(_gbadf.Images), i.Format)
	_bfd := _gbadf._dab.AddRelationship(_fbed, unioffice.ImageType)
	_gbadf.ContentTypes.EnsureDefault("\u0070\u006e\u0067", "\u0069m\u0061\u0067\u0065\u002f\u0070\u006eg")
	_gbadf.ContentTypes.EnsureDefault("\u006a\u0070\u0065\u0067", "\u0069\u006d\u0061\u0067\u0065\u002f\u006a\u0070\u0065\u0067")
	_gbadf.ContentTypes.EnsureDefault("\u006a\u0070\u0067", "\u0069\u006d\u0061\u0067\u0065\u002f\u006a\u0070\u0065\u0067")
	_gbadf.ContentTypes.EnsureDefault("\u0077\u006d\u0066", "i\u006d\u0061\u0067\u0065\u002f\u0078\u002d\u0077\u006d\u0066")
	_gbadf.ContentTypes.EnsureDefault(i.Format, "\u0069\u006d\u0061\u0067\u0065\u002f"+i.Format)
	_cea.SetRelID(_bfd.X().IdAttr)
	_cea.SetTarget(_fbed)
	return _cea, nil
}
func _fbee(_fgae io.ReaderAt, _fcfbg int64, _afe string) (*Document, error) {
	const _fdcd = "\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0052\u0065\u0061\u0064"
	if !license.GetLicenseKey().IsLicensed() && !_eece {
		fmt.Println("\u0055\u006e\u006ci\u0063\u0065\u006e\u0073e\u0064\u0020\u0076\u0065\u0072\u0073\u0069o\u006e\u0020\u006f\u0066\u0020\u0055\u006e\u0069\u004f\u0066\u0066\u0069\u0063\u0065")
		fmt.Println("\u002d\u0020\u0047e\u0074\u0020\u0061\u0020\u0074\u0072\u0069\u0061\u006c\u0020\u006c\u0069\u0063\u0065\u006e\u0073\u0065\u0020\u006f\u006e\u0020\u0068\u0074\u0074\u0070\u0073\u003a\u002f\u002fu\u006e\u0069\u0064\u006f\u0063\u002e\u0069\u006f")
		return nil, errors.New("\u0075\u006e\u0069\u006f\u0066\u0066\u0069\u0063\u0065\u0020\u006ci\u0063\u0065\u006e\u0073\u0065\u0020\u0072\u0065\u0071\u0075i\u0072\u0065\u0064")
	}
	_dgea := New()
	_dgea.Numbering._cbag = nil
	if len(_afe) > 0 {
		_dgea._feg = _afe
	} else {
		_gfeb, _ebd := license.GenRefId("\u0064\u0072")
		if _ebd != nil {
			logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _ebd)
			return nil, _ebd
		}
		_dgea._feg = _gfeb
	}
	if _dbd := license.Track(_dgea._feg, _fdcd); _dbd != nil {
		logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _dbd)
		return nil, _dbd
	}
	_ggabd, _dbff := tempstorage.TempDir("\u0075\u006e\u0069\u006f\u0066\u0066\u0069\u0063\u0065-\u0064\u006f\u0063\u0078")
	if _dbff != nil {
		return nil, _dbff
	}
	_dgea.TmpPath = _ggabd
	_efbe, _dbff := zip.NewReader(_fgae, _fcfbg)
	if _dbff != nil {
		return nil, fmt.Errorf("\u0070a\u0072s\u0069\u006e\u0067\u0020\u007a\u0069\u0070\u003a\u0020\u0025\u0073", _dbff)
	}
	_deg := []*zip.File{}
	_deg = append(_deg, _efbe.File...)
	_dcad := false
	for _, _eeeb := range _deg {
		if _eeeb.FileHeader.Name == "\u0064\u006f\u0063\u0050ro\u0070\u0073\u002f\u0063\u0075\u0073\u0074\u006f\u006d\u002e\u0078\u006d\u006c" {
			_dcad = true
			break
		}
	}
	if _dcad {
		_dgea.CreateCustomProperties()
	}
	_efda := _dgea.doc.ConformanceAttr
	_gacb := zippkg.DecodeMap{}
	_gacb.SetOnNewRelationshipFunc(_dgea.onNewRelationship)
	_gacb.AddTarget(unioffice.ContentTypesFilename, _dgea.ContentTypes.X(), "", 0)
	_gacb.AddTarget(unioffice.BaseRelsFilename, _dgea.Rels.X(), "", 0)
	if _dfcf := _gacb.Decode(_deg); _dfcf != nil {
		return nil, _dfcf
	}
	_dgea.doc.ConformanceAttr = _efda
	for _, _gcb := range _deg {
		if _gcb == nil {
			continue
		}
		if _cdfb := _dgea.AddExtraFileFromZip(_gcb); _cdfb != nil {
			return nil, _cdfb
		}
	}
	if _dcad {
		_eabe := false
		for _, _ddac := range _dgea.Rels.X().Relationship {
			if _ddac.TargetAttr == "\u0064\u006f\u0063\u0050ro\u0070\u0073\u002f\u0063\u0075\u0073\u0074\u006f\u006d\u002e\u0078\u006d\u006c" {
				_eabe = true
				break
			}
		}
		if !_eabe {
			_dgea.AddCustomRelationships()
		}
	}
	return _dgea, nil
}

// TableProperties returns the table style properties.
func (_ddceb Style) TableProperties() TableStyleProperties {
	if _ddceb._gaege.TblPr == nil {
		_ddceb._gaege.TblPr = wml.NewCT_TblPrBase()
	}
	return TableStyleProperties{_ddceb._gaege.TblPr}
}
func _aff(_ceg *wml.CT_TblWidth, _fbf float64) {
	_ceg.TypeAttr = wml.ST_TblWidthPct
	_ceg.WAttr = &wml.ST_MeasurementOrPercent{}
	_ceg.WAttr.ST_DecimalNumberOrPercent = &wml.ST_DecimalNumberOrPercent{}
	_ceg.WAttr.ST_DecimalNumberOrPercent.ST_UnqualifiedPercentage = unioffice.Int64(int64(_fbf * 50))
}

// IsChecked returns true if a FormFieldTypeCheckBox is checked.
func (_eecg FormField) IsChecked() bool {
	if _eecg._cbde.CheckBox == nil {
		return false
	}
	if _eecg._cbde.CheckBox.Checked != nil {
		return true
	}
	return false
}

// Caps returns true if run font is capitalized.
func (_dfbb RunProperties) Caps() bool { return _cadf(_dfbb._gbdb.Caps) }

// SetYOffset sets the Y offset for an image relative to the origin.
func (_gb AnchoredDrawing) SetYOffset(y measurement.Distance) {
	_gb._dgc.PositionV.Choice = &wml.WdCT_PosVChoice{}
	_gb._dgc.PositionV.Choice.PosOffset = unioffice.Int32(int32(y / measurement.EMU))
}

// SetLeftPct sets the cell left margin
func (_gca CellMargins) SetLeftPct(pct float64) {
	_gca._cdae.Left = wml.NewCT_TblWidth()
	_aff(_gca._cdae.Left, pct)
}

// Header is a header for a document section.
type Header struct {
	_dbagd *Document
	_deae  *wml.Hdr
}

// AddParagraph adds a paragraph to the table cell.
func (_aed Cell) AddParagraph() Paragraph {
	_fae := wml.NewEG_BlockLevelElts()
	_aed._gge.EG_BlockLevelElts = append(_aed._gge.EG_BlockLevelElts, _fae)
	_gdg := wml.NewEG_ContentBlockContent()
	_fae.EG_ContentBlockContent = append(_fae.EG_ContentBlockContent, _gdg)
	_dgaa := wml.NewCT_P()
	_gdg.P = append(_gdg.P, _dgaa)
	return Paragraph{_aed._dga, _dgaa}
}

// X returns the inner wrapped XML type.
func (_gfadf Run) X() *wml.CT_R { return _gfadf._adaad }
func _bfgge(_ccee *wml.CT_P, _dadd map[string]string) {
	for _, _ccdd := range _ccee.EG_PContent {
		for _, _ddcae := range _ccdd.EG_ContentRunContent {
			if _ddcae.R != nil {
				for _, _ece := range _ddcae.R.EG_RunInnerContent {
					_adfd := _ece.Drawing
					if _adfd != nil {
						for _, _feee := range _adfd.Anchor {
							for _, _gfae := range _feee.Graphic.GraphicData.Any {
								switch _gebc := _gfae.(type) {
								case *picture.Pic:
									if _gebc.BlipFill != nil && _gebc.BlipFill.Blip != nil {
										_acgc(_gebc.BlipFill.Blip, _dadd)
									}
								default:
								}
							}
						}
						for _, _gfbg := range _adfd.Inline {
							for _, _eggb := range _gfbg.Graphic.GraphicData.Any {
								switch _eefc := _eggb.(type) {
								case *picture.Pic:
									if _eefc.BlipFill != nil && _eefc.BlipFill.Blip != nil {
										_acgc(_eefc.BlipFill.Blip, _dadd)
									}
								default:
								}
							}
						}
					}
				}
			}
		}
	}
}

// Section is the beginning of a new section.
type Section struct {
	_afafb *Document
	_ddcag *wml.CT_SectPr
}

// Styles is the document wide styles contained in styles.xml.
type Styles struct{ _abca *wml.Styles }

// SetChecked marks a FormFieldTypeCheckBox as checked or unchecked.
func (_ecef FormField) SetChecked(b bool) {
	if _ecef._cbde.CheckBox == nil {
		return
	}
	if !b {
		_ecef._cbde.CheckBox.Checked = nil
	} else {
		_ecef._cbde.CheckBox.Checked = wml.NewCT_OnOff()
	}
}

// Tables returns the tables defined in the header.
func (_gfbd Header) Tables() []Table {
	_fbgf := []Table{}
	if _gfbd._deae == nil {
		return nil
	}
	for _, _gfcf := range _gfbd._deae.EG_ContentBlockContent {
		for _, _efgc := range _gfbd._dbagd.tables(_gfcf) {
			_fbgf = append(_fbgf, _efgc)
		}
	}
	return _fbgf
}

// SetHeight allows controlling the height of a row within a table.
func (_ccgc RowProperties) SetHeight(ht measurement.Distance, rule wml.ST_HeightRule) {
	if rule == wml.ST_HeightRuleUnset {
		_ccgc._acgb.TrHeight = nil
	} else {
		_cabe := wml.NewCT_Height()
		_cabe.HRuleAttr = rule
		_cabe.ValAttr = &sharedTypes.ST_TwipsMeasure{}
		_cabe.ValAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(ht / measurement.Twips))
		_ccgc._acgb.TrHeight = []*wml.CT_Height{_cabe}
	}
}

// // SetBeforeLineSpacing sets spacing above paragraph in line units.
func (_dccgc Paragraph) SetBeforeLineSpacing(d measurement.Distance) {
	_dccgc.ensurePPr()
	if _dccgc._eagd.PPr.Spacing == nil {
		_dccgc._eagd.PPr.Spacing = wml.NewCT_Spacing()
	}
	_bcdaf := _dccgc._eagd.PPr.Spacing
	_bcdaf.BeforeLinesAttr = unioffice.Int64(int64(d / measurement.Twips))
}

// SetMultiLevelType sets the multilevel type.
func (_caeff NumberingDefinition) SetMultiLevelType(t wml.ST_MultiLevelType) {
	if t == wml.ST_MultiLevelTypeUnset {
		_caeff._agff.MultiLevelType = nil
	} else {
		_caeff._agff.MultiLevelType = wml.NewCT_MultiLevelType()
		_caeff._agff.MultiLevelType.ValAttr = t
	}
}

// X returns the inner wrapped XML type.
func (_eaadf Footnote) X() *wml.CT_FtnEdn { return _eaadf._bgcda }

// EastAsiaFont returns the name of run font family for East Asia.
func (_ggbe RunProperties) EastAsiaFont() string {
	if _efgg := _ggbe._gbdb.RFonts; _efgg != nil {
		if _efgg.EastAsiaAttr != nil {
			return *_efgg.EastAsiaAttr
		}
	}
	return ""
}

// TableStyleProperties are table properties as defined in a style.
type TableStyleProperties struct{ _degc *wml.CT_TblPrBase }

// SetBottom sets the bottom border to a specified type, color and thickness.
func (_aeccb ParagraphBorders) SetBottom(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_aeccb._fdge.Bottom = wml.NewCT_Border()
	_bbgf(_aeccb._fdge.Bottom, t, c, thickness)
}

// MailMerge finds mail merge fields and replaces them with the text provided.  It also removes
// the mail merge source info from the document settings.
func (_fdbce *Document) MailMerge(mergeContent map[string]string) {
	_ebag := _fdbce.mergeFields()
	_bcfa := map[Paragraph][]Run{}
	for _, _fada := range _ebag {
		_gdgg, _geefe := mergeContent[_fada._gdfge]
		if _geefe {
			if _fada._dbfgc {
				_gdgg = strings.ToUpper(_gdgg)
			} else if _fada._fgaa {
				_gdgg = strings.ToLower(_gdgg)
			} else if _fada._bcdae {
				_gdgg = strings.Title(_gdgg)
			} else if _fada._bddb {
				_agggde := bytes.Buffer{}
				for _egge, _ccaf := range _gdgg {
					if _egge == 0 {
						_agggde.WriteRune(unicode.ToUpper(_ccaf))
					} else {
						_agggde.WriteRune(_ccaf)
					}
				}
				_gdgg = _agggde.String()
			}
			if _gdgg != "" && _fada._dfga != "" {
				_gdgg = _fada._dfga + _gdgg
			}
			if _gdgg != "" && _fada._cgec != "" {
				_gdgg = _gdgg + _fada._cgec
			}
		}
		if _fada._debdg {
			if len(_fada._ceaf.FldSimple) == 1 && len(_fada._ceaf.FldSimple[0].EG_PContent) == 1 && len(_fada._ceaf.FldSimple[0].EG_PContent[0].EG_ContentRunContent) == 1 {
				_bffe := &wml.EG_ContentRunContent{}
				_bffe.R = _fada._ceaf.FldSimple[0].EG_PContent[0].EG_ContentRunContent[0].R
				_fada._ceaf.FldSimple = nil
				_fffcc := Run{_fdbce, _bffe.R}
				_fffcc.ClearContent()
				_fffcc.AddText(_gdgg)
				_fada._ceaf.EG_ContentRunContent = append(_fada._ceaf.EG_ContentRunContent, _bffe)
			}
		} else {
			_baeg := _fada._abdbd.Runs()
			for _ccff := _fada._bbcb; _ccff <= _fada._gfaf; _ccff++ {
				if _ccff == _fada._cdcbe+1 {
					_baeg[_ccff].ClearContent()
					_baeg[_ccff].AddText(_gdgg)
				} else {
					_bcfa[_fada._abdbd] = append(_bcfa[_fada._abdbd], _baeg[_ccff])
				}
			}
		}
	}
	for _gaaa, _eegb := range _bcfa {
		for _, _agbfa := range _eegb {
			_gaaa.RemoveRun(_agbfa)
		}
	}
	_fdbce.Settings.RemoveMailMerge()
}

// SetNumberingLevel sets the numbering level of a paragraph.  If used, then the
// NumberingDefinition must also be set via SetNumberingDefinition or
// SetNumberingDefinitionByID.
func (_geedg Paragraph) SetNumberingLevel(listLevel int) {
	_geedg.ensurePPr()
	if _geedg._eagd.PPr.NumPr == nil {
		_geedg._eagd.PPr.NumPr = wml.NewCT_NumPr()
	}
	_ebee := wml.NewCT_DecimalNumber()
	_ebee.ValAttr = int64(listLevel)
	_geedg._eagd.PPr.NumPr.Ilvl = _ebee
}

// SetStart sets the cell start margin
func (_egb CellMargins) SetStart(d measurement.Distance) {
	_egb._cdae.Start = wml.NewCT_TblWidth()
	_age(_egb._cdae.Start, d)
}

// FontTable return document fontTable element.
func (_cada *Document) FontTable() *wml.Fonts { return _cada._eaa }
func _gcdg(_cecf *wml.CT_Tbl, _gfebc map[string]string) {
	for _, _dfaa := range _cecf.EG_ContentRowContent {
		for _, _egdcc := range _dfaa.Tr {
			for _, _daff := range _egdcc.EG_ContentCellContent {
				for _, _gddb := range _daff.Tc {
					for _, _ebab := range _gddb.EG_BlockLevelElts {
						for _, _afef := range _ebab.EG_ContentBlockContent {
							for _, _edgee := range _afef.P {
								_cbdfg(_edgee, _gfebc)
							}
							for _, _bfaf := range _afef.Tbl {
								_gcdg(_bfaf, _gfebc)
							}
						}
					}
				}
			}
		}
	}
}

// X returns the inner wrapped XML type.
func (_edgf *Document) X() *wml.Document { return _edgf.doc }

// SetSize sets the font size for a run.
func (_ggcbb RunProperties) SetSize(size measurement.Distance) {
	_ggcbb._gbdb.Sz = wml.NewCT_HpsMeasure()
	_ggcbb._gbdb.Sz.ValAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(size / measurement.HalfPoint))
	_ggcbb._gbdb.SzCs = wml.NewCT_HpsMeasure()
	_ggcbb._gbdb.SzCs.ValAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(size / measurement.HalfPoint))
}

// SaveToFile writes the document out to a file.
func (_aeed *Document) SaveToFile(path string) error {
	_dacd, _cdgg := os.Create(path)
	if _cdgg != nil {
		return _cdgg
	}
	defer _dacd.Close()
	return _aeed.Save(_dacd)
}

// SetNextStyle sets the style that the next paragraph will use.
func (_egeb Style) SetNextStyle(name string) {
	if name == "" {
		_egeb._gaege.Next = nil
	} else {
		_egeb._gaege.Next = wml.NewCT_String()
		_egeb._gaege.Next.ValAttr = name
	}
}

// New constructs an empty document that content can be added to.
func New() *Document {
	_agf := &Document{doc: wml.NewDocument()}
	_agf.ContentTypes = common.NewContentTypes()
	_agf.doc.Body = wml.NewCT_Body()
	_agf.doc.ConformanceAttr = sharedTypes.ST_ConformanceClassTransitional
	_agf._dab = common.NewRelationships()
	_agf.AppProperties = common.NewAppProperties()
	_agf.CoreProperties = common.NewCoreProperties()
	_agf.ContentTypes.AddOverride("\u002fw\u006fr\u0064\u002f\u0064\u006f\u0063u\u006d\u0065n\u0074\u002e\u0078\u006d\u006c", "\u0061p\u0070\u006c\u0069c\u0061\u0074\u0069o\u006e/v\u006e\u0064\u002e\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072\u006d\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072dp\u0072o\u0063\u0065\u0073\u0073\u0069\u006eg\u006d\u006c\u002e\u0064\u006fc\u0075\u006d\u0065\u006e\u0074\u002e\u006d\u0061\u0069\u006e\u002bx\u006d\u006c")
	_agf.Settings = NewSettings()
	_agf._dab.AddRelationship("\u0073\u0065\u0074t\u0069\u006e\u0067\u0073\u002e\u0078\u006d\u006c", unioffice.SettingsType)
	_agf.ContentTypes.AddOverride("\u002fw\u006fr\u0064\u002f\u0073\u0065\u0074t\u0069\u006eg\u0073\u002e\u0078\u006d\u006c", "\u0061\u0070\u0070\u006c\u0069\u0063\u0061\u0074\u0069o\u006e\u002fv\u006e\u0064\u002e\u006f\u0070\u0065\u006e\u0078\u006dl\u0066\u006f\u0072\u006da\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c.\u0073\u0065\u0074\u0074\u0069\u006e\u0067\u0073\u002b\u0078\u006d\u006c")
	_agf.Rels = common.NewRelationships()
	_agf.Rels.AddRelationship(unioffice.RelativeFilename(unioffice.DocTypeDocument, "", unioffice.CorePropertiesType, 0), unioffice.CorePropertiesType)
	_agf.Rels.AddRelationship("\u0064\u006fc\u0050\u0072\u006fp\u0073\u002f\u0061\u0070\u0070\u002e\u0078\u006d\u006c", unioffice.ExtendedPropertiesType)
	_agf.Rels.AddRelationship("\u0077\u006f\u0072\u0064\u002f\u0064\u006f\u0063\u0075\u006d\u0065\u006et\u002e\u0078\u006d\u006c", unioffice.OfficeDocumentType)
	_agf.Numbering = NewNumbering()
	_agf.Numbering.InitializeDefault()
	_agf.ContentTypes.AddOverride("\u002f\u0077\u006f\u0072d/\u006e\u0075\u006d\u0062\u0065\u0072\u0069\u006e\u0067\u002e\u0078\u006d\u006c", "\u0061\u0070\u0070\u006c\u0069c\u0061\u0074\u0069\u006f\u006e\u002f\u0076n\u0064\u002e\u006f\u0070\u0065\u006e\u0078\u006d\u006c\u0066\u006f\u0072\u006d\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063e\u0073\u0073\u0069\u006e\u0067\u006d\u006c\u002e\u006e\u0075\u006d\u0062e\u0072\u0069\u006e\u0067\u002b\u0078m\u006c")
	_agf._dab.AddRelationship("\u006e\u0075\u006d\u0062\u0065\u0072\u0069\u006e\u0067\u002e\u0078\u006d\u006c", unioffice.NumberingType)
	_agf.Styles = NewStyles()
	_agf.Styles.InitializeDefault()
	_agf.ContentTypes.AddOverride("\u002f\u0077o\u0072\u0064\u002fs\u0074\u0079\u006c\u0065\u0073\u002e\u0078\u006d\u006c", "\u0061p\u0070l\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0076\u006e\u0064.\u006f\u0070\u0065\u006ex\u006d\u006c\u0066\u006f\u0072m\u0061\u0074\u0073\u002d\u006f\u0066\u0066\u0069\u0063\u0065\u0064\u006f\u0063\u0075\u006d\u0065\u006e\u0074\u002e\u0077\u006f\u0072\u0064\u0070\u0072\u006f\u0063\u0065\u0073\u0073\u0069n\u0067\u006d\u006c\u002e\u0073\u0074\u0079\u006ce\u0073\u002b\u0078\u006d\u006c")
	_agf._dab.AddRelationship("\u0073\u0074\u0079\u006c\u0065\u0073\u002e\u0078\u006d\u006c", unioffice.StylesType)
	_agf.doc.Body = wml.NewCT_Body()
	return _agf
}

// Caps returns true if paragraph font is capitalized.
func (_eebaf ParagraphProperties) Caps() bool { return _cadf(_eebaf._dfaf.RPr.Caps) }

// Color controls the run or styles color.
type Color struct{ _ec *wml.CT_Color }

// NewWatermarkPicture generates new WatermarkPicture.
func NewWatermarkPicture() WatermarkPicture {
	_eceg := vml.NewShapetype()
	_fadd := vml.NewEG_ShapeElements()
	_fadd.Formulas = _edbf()
	_fadd.Path = _dfedg()
	_fadd.Lock = _bgbf()
	_eceg.EG_ShapeElements = []*vml.EG_ShapeElements{_fadd}
	var (
		_afeg  = "\u005f\u0078\u0030\u0030\u0030\u0030\u005f\u0074\u0037\u0035"
		_bgfab = "2\u0031\u0036\u0030\u0030\u002c\u0032\u0031\u0036\u0030\u0030"
		_gbeeb = float32(75.0)
		_bgebf = "\u006d\u0040\u0034\u00405l\u0040\u0034\u0040\u0031\u0031\u0040\u0039\u0040\u0031\u0031\u0040\u0039\u0040\u0035x\u0065"
	)
	_eceg.IdAttr = &_afeg
	_eceg.CoordsizeAttr = &_bgfab
	_eceg.SptAttr = &_gbeeb
	_eceg.PreferrelativeAttr = sharedTypes.ST_TrueFalseTrue
	_eceg.PathAttr = &_bgebf
	_eceg.FilledAttr = sharedTypes.ST_TrueFalseFalse
	_eceg.StrokedAttr = sharedTypes.ST_TrueFalseFalse
	_aebdg := vml.NewShape()
	_deebg := vml.NewEG_ShapeElements()
	_deebg.Imagedata = _ebbad()
	_aebdg.EG_ShapeElements = []*vml.EG_ShapeElements{_deebg}
	var (
		_babdaf = "\u0057\u006f\u0072\u0064\u0050\u0069\u0063\u0074\u0075\u0072e\u0057\u0061\u0074\u0065\u0072\u006d\u0061r\u006b\u0031\u0036\u0033\u0032\u0033\u0031\u0036\u0035\u0039\u0035"
		_baebe  = "\u005f\u0078\u00300\u0030\u0030\u005f\u0073\u0032\u0030\u0035\u0031"
		_cdfac  = "#\u005f\u0078\u0030\u0030\u0030\u0030\u005f\u0074\u0037\u0035"
		_edfc   = ""
		_cfgb   = "\u0070os\u0069t\u0069o\u006e\u003a\u0061\u0062\u0073\u006fl\u0075\u0074\u0065\u003bm\u0061\u0072\u0067\u0069\u006e\u002d\u006c\u0065\u0066\u0074\u003a\u0030\u003bma\u0072\u0067\u0069\u006e\u002d\u0074\u006f\u0070\u003a\u0030\u003b\u0077\u0069\u0064\u0074\u0068\u003a\u0030\u0070\u0074;\u0068e\u0069\u0067\u0068\u0074\u003a\u0030\u0070\u0074\u003b\u007a\u002d\u0069\u006ed\u0065\u0078:\u002d\u0032\u00351\u0036\u0035\u0038\u0032\u0034\u0030\u003b\u006d\u0073o-\u0070\u006f\u0073i\u0074\u0069\u006f\u006e-\u0068\u006f\u0072\u0069\u007a\u006fn\u0074\u0061l\u003a\u0063\u0065\u006e\u0074\u0065\u0072\u003bm\u0073\u006f\u002d\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e\u002d\u0068\u006f\u0072\u0069\u007a\u006f\u006e\u0074\u0061\u006c\u002drela\u0074\u0069\u0076\u0065\u003a\u006d\u0061\u0072\u0067\u0069\u006e\u003b\u006d\u0073\u006f\u002d\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e\u002d\u0076\u0065\u0072t\u0069c\u0061l\u003a\u0063\u0065\u006e\u0074\u0065\u0072\u003b\u006d\u0073\u006f\u002d\u0070\u006f\u0073\u0069\u0074\u0069\u006f\u006e-\u0076\u0065r\u0074\u0069c\u0061l\u002d\u0072\u0065\u006c\u0061\u0074i\u0076\u0065\u003a\u006d\u0061\u0072\u0067\u0069\u006e"
	)
	_aebdg.IdAttr = &_babdaf
	_aebdg.SpidAttr = &_baebe
	_aebdg.TypeAttr = &_cdfac
	_aebdg.AltAttr = &_edfc
	_aebdg.StyleAttr = &_cfgb
	_aebdg.AllowincellAttr = sharedTypes.ST_TrueFalseFalse
	_eggeb := wml.NewCT_Picture()
	_eggeb.Any = []unioffice.Any{_eceg, _aebdg}
	return WatermarkPicture{_cdff: _eggeb, _fdgfa: _aebdg, _acbd: _eceg}
}

// X returns the internally wrapped *wml.CT_SectPr.
func (_fcgb Section) X() *wml.CT_SectPr { return _fcgb._ddcag }

// SetAlignment set alignment of paragraph.
func (_gfdf Paragraph) SetAlignment(alignment wml.ST_Jc) {
	_gfdf.ensurePPr()
	if _gfdf._eagd.PPr.Jc == nil {
		_gfdf._eagd.PPr.Jc = wml.NewCT_Jc()
	}
	_gfdf._eagd.PPr.Jc.ValAttr = alignment
}
func _feadc(_fabd *wml.CT_Border, _fgfdc wml.ST_Border, _geabd color.Color, _dgbga measurement.Distance) {
	_fabd.ValAttr = _fgfdc
	_fabd.ColorAttr = &wml.ST_HexColor{}
	if _geabd.IsAuto() {
		_fabd.ColorAttr.ST_HexColorAuto = wml.ST_HexColorAutoAuto
	} else {
		_fabd.ColorAttr.ST_HexColorRGB = _geabd.AsRGBString()
	}
	if _dgbga != measurement.Zero {
		_fabd.SzAttr = unioffice.Uint64(uint64(_dgbga / measurement.Point * 8))
	}
}

// InsertParagraphBefore adds a new empty paragraph before the relativeTo
// paragraph.
func (_aggfc *Document) InsertParagraphBefore(relativeTo Paragraph) Paragraph {
	return _aggfc.insertParagraph(relativeTo, true)
}

// Bold returns true if run font is bold.
func (_caega RunProperties) Bold() bool {
	_abgd := _caega._gbdb
	return _cadf(_abgd.B) || _cadf(_abgd.BCs)
}

// Save writes the document to an io.Writer in the Zip package format.
func (_cec *Document) Save(w io.Writer) error { return _cec.save(w, _cec._feg) }
func _cbdfg(_cagc *wml.CT_P, _deeb map[string]string) {
	for _, _deebe := range _cagc.EG_PContent {
		if _deebe.Hyperlink != nil && _deebe.Hyperlink.IdAttr != nil {
			if _edbd, _ddgag := _deeb[*_deebe.Hyperlink.IdAttr]; _ddgag {
				*_deebe.Hyperlink.IdAttr = _edbd
			}
		}
	}
}

// PutNodeBefore put node to position before relativeTo.
func (_acag *Document) PutNodeBefore(relativeTo, node Node) { _acag.putNode(relativeTo, node, true) }

// SetTextStyleBold set text style of watermark to bold.
func (_cgbf *WatermarkText) SetTextStyleBold(value bool) {
	if _cgbf._bfbf != nil {
		_feeg := _cgbf.GetStyle()
		_feeg.SetBold(value)
		_cgbf.SetStyle(_feeg)
	}
}

// SetKeepNext controls if the paragraph is kept with the next paragraph.
func (_aebf ParagraphStyleProperties) SetKeepNext(b bool) {
	if !b {
		_aebf._gfee.KeepNext = nil
	} else {
		_aebf._gfee.KeepNext = wml.NewCT_OnOff()
	}
}

// RemoveRun removes a child run from a paragraph.
func (_ggdag Paragraph) RemoveRun(r Run) {
	for _, _aabc := range _ggdag._eagd.EG_PContent {
		for _cdecd, _ddece := range _aabc.EG_ContentRunContent {
			if _ddece.R == r._adaad {
				copy(_aabc.EG_ContentRunContent[_cdecd:], _aabc.EG_ContentRunContent[_cdecd+1:])
				_aabc.EG_ContentRunContent = _aabc.EG_ContentRunContent[0 : len(_aabc.EG_ContentRunContent)-1]
			}
			if _ddece.Sdt != nil && _ddece.Sdt.SdtContent != nil {
				for _cgee, _abbd := range _ddece.Sdt.SdtContent.EG_ContentRunContent {
					if _abbd.R == r._adaad {
						copy(_ddece.Sdt.SdtContent.EG_ContentRunContent[_cgee:], _ddece.Sdt.SdtContent.EG_ContentRunContent[_cgee+1:])
						_ddece.Sdt.SdtContent.EG_ContentRunContent = _ddece.Sdt.SdtContent.EG_ContentRunContent[0 : len(_ddece.Sdt.SdtContent.EG_ContentRunContent)-1]
					}
				}
			}
		}
	}
}

// SetLastRow controls the conditional formatting for the last row in a table.
// This is called the 'Total' row within Word.
func (_cfdgc TableLook) SetLastRow(on bool) {
	if !on {
		_cfdgc.ctTblLook.LastRowAttr = &sharedTypes.ST_OnOff{}
		_cfdgc.ctTblLook.LastRowAttr.ST_OnOff1 = sharedTypes.ST_OnOff1Off
	} else {
		_cfdgc.ctTblLook.LastRowAttr = &sharedTypes.ST_OnOff{}
		_cfdgc.ctTblLook.LastRowAttr.ST_OnOff1 = sharedTypes.ST_OnOff1On
	}
}

// GetNumberingLevelByIds returns a NumberingLevel by its NumId and LevelId attributes
// or an empty one if not found.
func (_gedc *Document) GetNumberingLevelByIds(numId, levelId int64) NumberingLevel {
	for _, _geab := range _gedc.Numbering._cbag.Num {
		if _geab != nil && _geab.NumIdAttr == numId {
			_gfgfb := _geab.AbstractNumId.ValAttr
			for _, _acec := range _gedc.Numbering._cbag.AbstractNum {
				if _acec.AbstractNumIdAttr == _gfgfb {
					if _acec.NumStyleLink != nil && len(_acec.Lvl) == 0 {
						if _gfgfd, _dcdb := _gedc.Styles.SearchStyleById(_acec.NumStyleLink.ValAttr); _dcdb {
							if _gfgfd.ParagraphProperties().NumId() > -1 {
								return _gedc.GetNumberingLevelByIds(_gfgfd.ParagraphProperties().NumId(), levelId)
							}
						}
					}
					for _, _afgb := range _acec.Lvl {
						if _afgb.IlvlAttr == levelId {
							return NumberingLevel{_afgb}
						}
					}
				}
			}
		}
	}
	return NumberingLevel{}
}

// SetFooter sets a section footer.
func (_cafe Section) SetFooter(f Footer, t wml.ST_HdrFtr) {
	_cecfe := wml.NewEG_HdrFtrReferences()
	_cafe._ddcag.EG_HdrFtrReferences = append(_cafe._ddcag.EG_HdrFtrReferences, _cecfe)
	_cecfe.FooterReference = wml.NewCT_HdrFtrRef()
	_cecfe.FooterReference.TypeAttr = t
	_ggdff := _cafe._afafb._dab.FindRIDForN(f.Index(), unioffice.FooterType)
	if _ggdff == "" {
		logger.Log.Debug("\u0075\u006ea\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0064\u0065\u0074\u0065\u0072\u006d\u0069\u006e\u0065\u0020\u0066\u006f\u006f\u0074\u0065r \u0049\u0044")
	}
	_cecfe.FooterReference.IdAttr = _ggdff
}

// AddLevel adds a new numbering level to a NumberingDefinition.
func (_bcbf NumberingDefinition) AddLevel() NumberingLevel {
	_fdgf := wml.NewCT_Lvl()
	_fdgf.Start = &wml.CT_DecimalNumber{ValAttr: 1}
	_fdgf.IlvlAttr = int64(len(_bcbf._agff.Lvl))
	_bcbf._agff.Lvl = append(_bcbf._agff.Lvl, _fdgf)
	return NumberingLevel{_fdgf}
}

// Shadow returns true if run shadow is on.
func (_fdff RunProperties) Shadow() bool { return _cadf(_fdff._gbdb.Shadow) }

// InsertRunAfter inserts a run in the paragraph after the relative run.
func (_ggeff Paragraph) InsertRunAfter(relativeTo Run) Run {
	return _ggeff.insertRun(relativeTo, false)
}

// SetOffset sets the offset of the image relative to the origin, which by
// default this is the top-left corner of the page. Offset is incompatible with
// SetAlignment, whichever is called last is applied.
func (_eed AnchoredDrawing) SetOffset(x, y measurement.Distance) {
	_eed.SetXOffset(x)
	_eed.SetYOffset(y)
}

// SetSmallCaps sets the run to small caps.
func (_dabc RunProperties) SetSmallCaps(b bool) {
	if !b {
		_dabc._gbdb.SmallCaps = nil
	} else {
		_dabc._gbdb.SmallCaps = wml.NewCT_OnOff()
	}
}

// SetTop sets the top border to a specified type, color and thickness.
func (_agge ParagraphBorders) SetTop(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_agge._fdge.Top = wml.NewCT_Border()
	_bbgf(_agge._fdge.Top, t, c, thickness)
}

// SetHeader sets a section header.
func (_fcbbb Section) SetHeader(h Header, t wml.ST_HdrFtr) {
	_ggeg := wml.NewEG_HdrFtrReferences()
	_fcbbb._ddcag.EG_HdrFtrReferences = append(_fcbbb._ddcag.EG_HdrFtrReferences, _ggeg)
	_ggeg.HeaderReference = wml.NewCT_HdrFtrRef()
	_ggeg.HeaderReference.TypeAttr = t
	_bbeff := _fcbbb._afafb._dab.FindRIDForN(h.Index(), unioffice.HeaderType)
	if _bbeff == "" {
		logger.Log.Debug("\u0075\u006ea\u0062\u006c\u0065\u0020\u0074\u006f\u0020\u0064\u0065\u0074\u0065\u0072\u006d\u0069\u006e\u0065\u0020\u0068\u0065\u0061\u0064\u0065r \u0049\u0044")
	}
	_ggeg.HeaderReference.IdAttr = _bbeff
}

// MergeFields returns the list of all mail merge fields found in the document.
func (_adbd Document) MergeFields() []string {
	_edab := map[string]struct{}{}
	for _, _abgga := range _adbd.mergeFields() {
		_edab[_abgga._gdfge] = struct{}{}
	}
	_daddc := []string{}
	for _dddg := range _edab {
		_daddc = append(_daddc, _dddg)
	}
	return _daddc
}
func _caac(_fabgb *Document) map[int64]map[int64]int64 {
	_daed := _fabgb.Paragraphs()
	_baag := make(map[int64]map[int64]int64, 0)
	for _, _eebaa := range _daed {
		_aebga := _cbcf(_fabgb, _eebaa)
		if _aebga.NumberingLevel != nil && _aebga.AbstractNumId != nil {
			_cddc := *_aebga.AbstractNumId
			if _, _fggg := _baag[_cddc]; _fggg {
				if _aadg := _aebga.NumberingLevel.X(); _aadg != nil {
					if _, _faaa := _baag[_cddc][_aadg.IlvlAttr]; _faaa {
						_baag[_cddc][_aadg.IlvlAttr]++
					} else {
						_baag[_cddc][_aadg.IlvlAttr] = 1
					}
				}
			} else {
				if _agbbd := _aebga.NumberingLevel.X(); _agbbd != nil {
					_baag[_cddc] = map[int64]int64{_agbbd.IlvlAttr: 1}
				}
			}
		}
	}
	return _baag
}

// CellMargins are the margins for an individual cell.
type CellMargins struct{ _cdae *wml.CT_TcMar }

// X returns the inner wrapped XML type.
func (_ebaab ParagraphProperties) X() *wml.CT_PPr { return _ebaab._dfaf }

// GetImageByRelID returns an ImageRef with the associated relation ID in the
// document.
func (_fgde *Document) GetImageByRelID(relID string) (common.ImageRef, bool) {
	_daf := _fgde._dab.GetTargetByRelId(relID)
	_cbca := ""
	for _, _feefe := range _fgde._cbfd {
		if _cbca != "" {
			break
		}
		_cbca = _feefe.GetTargetByRelId(relID)
	}
	for _, _fgad := range _fgde.Images {
		if _fgad.RelID() == relID {
			return _fgad, true
		}
		if _daf != "" {
			_bab := strings.Replace(_fgad.Target(), "\u0077\u006f\u0072d\u002f", "", 1)
			if _bab == _daf {
				if _fgad.RelID() == "" {
					_fgad.SetRelID(relID)
				}
				return _fgad, true
			}
		}
		if _cbca != "" {
			_bdfb := strings.Replace(_fgad.Target(), "\u0077\u006f\u0072d\u002f", "", 1)
			if _bdfb == _cbca {
				if _fgad.RelID() == "" {
					_fgad.SetRelID(relID)
				}
				return _fgad, true
			}
		}
	}
	return common.ImageRef{}, false
}

const (
	FormFieldTypeUnknown FormFieldType = iota
	FormFieldTypeText
	FormFieldTypeCheckBox
	FormFieldTypeDropDown
)

func _beaea(_gfcgf *Document, _aegc []*wml.EG_ContentBlockContent, _ggceg *TableInfo) []Node {
	_ffbf := []Node{}
	for _, _dgfg := range _aegc {
		if _gbeb := _dgfg.Sdt; _gbeb != nil {
			if _ffea := _gbeb.SdtContent; _ffea != nil {
				_ffbf = append(_ffbf, _ecdc(_gfcgf, _ffea.P, _ggceg, nil)...)
			}
		}
		_ffbf = append(_ffbf, _ecdc(_gfcgf, _dgfg.P, _ggceg, nil)...)
		for _, _cccc := range _dgfg.Tbl {
			_acffc := Table{_gfcgf, _cccc}
			_bcbb, _ := _gfcgf.Styles.SearchStyleById(_acffc.Style())
			_afdeb := []Node{}
			for _faae, _befg := range _cccc.EG_ContentRowContent {
				for _, _bagg := range _befg.Tr {
					for _agee, _edabe := range _bagg.EG_ContentCellContent {
						for _, _aefcb := range _edabe.Tc {
							_acdb := &TableInfo{Table: _cccc, Row: _bagg, Cell: _aefcb, RowIndex: _faae, ColIndex: _agee}
							for _, _daeg := range _aefcb.EG_BlockLevelElts {
								_afdeb = append(_afdeb, _beaea(_gfcgf, _daeg.EG_ContentBlockContent, _acdb)...)
							}
						}
					}
				}
			}
			_ffbf = append(_ffbf, Node{_cdbd: _gfcgf, _ggda: &_acffc, Style: _bcbb, Children: _afdeb})
		}
	}
	return _ffbf
}

// StructuredDocumentTag are a tagged bit of content in a document.
type StructuredDocumentTag struct {
	_fdad  *Document
	_afadb *wml.CT_SdtBlock
}

// FindNodeByCondition return node based on condition function,
// if wholeElements is true, its will extract childs as next node elements.
func (_fgcc *Nodes) FindNodeByCondition(f func(_cffga *Node) bool, wholeElements bool) []Node {
	_dffe := []Node{}
	for _, _eccg := range _fgcc._gabfc {
		if f(&_eccg) {
			_dffe = append(_dffe, _eccg)
		}
		if wholeElements {
			_acff := Nodes{_gabfc: _eccg.Children}
			_dffe = append(_dffe, _acff.FindNodeByCondition(f, wholeElements)...)
		}
	}
	return _dffe
}

// AddParagraph adds a paragraph to the footer.
func (_ddecd Footer) AddParagraph() Paragraph {
	_bdaf := wml.NewEG_ContentBlockContent()
	_ddecd._fcc.EG_ContentBlockContent = append(_ddecd._fcc.EG_ContentBlockContent, _bdaf)
	_eagef := wml.NewCT_P()
	_bdaf.P = append(_bdaf.P, _eagef)
	return Paragraph{_ddecd._aegg, _eagef}
}

// SetWrapPathStart sets wrapPath start value.
func (_ffgc AnchorDrawWrapOptions) SetWrapPathStart(coordinate *dml.CT_Point2D) {
	_ffgc._dd = coordinate
}

// Properties returns the row properties.
func (_bacg Row) Properties() RowProperties {
	if _bacg.ctRow.TrPr == nil {
		_bacg.ctRow.TrPr = wml.NewCT_TrPr()
	}
	return RowProperties{_bacg.ctRow.TrPr}
}

// SetEmboss sets the run to embossed text.
func (_dbaefc RunProperties) SetEmboss(b bool) {
	if !b {
		_dbaefc._gbdb.Emboss = nil
	} else {
		_dbaefc._gbdb.Emboss = wml.NewCT_OnOff()
	}
}

// SetAlignment controls the paragraph alignment
func (_gbdg ParagraphProperties) SetAlignment(align wml.ST_Jc) {
	if align == wml.ST_JcUnset {
		_gbdg._dfaf.Jc = nil
	} else {
		_gbdg._dfaf.Jc = wml.NewCT_Jc()
		_gbdg._dfaf.Jc.ValAttr = align
	}
}

// Properties returns the cell properties.
func (_aaad Cell) Properties() CellProperties {
	if _aaad._gge.TcPr == nil {
		_aaad._gge.TcPr = wml.NewCT_TcPr()
	}
	return CellProperties{_aaad._gge.TcPr}
}

// SetEndIndent controls the end indentation.
func (_geaf ParagraphProperties) SetEndIndent(m measurement.Distance) {
	if _geaf._dfaf.Ind == nil {
		_geaf._dfaf.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_geaf._dfaf.Ind.EndAttr = nil
	} else {
		_geaf._dfaf.Ind.EndAttr = &wml.ST_SignedTwipsMeasure{}
		_geaf._dfaf.Ind.EndAttr.Int64 = unioffice.Int64(int64(m / measurement.Twips))
	}
}

// UnderlineColor returns the hex color value of run underline.
func (_eagdf RunProperties) UnderlineColor() string {
	if _fccf := _eagdf._gbdb.U; _fccf != nil {
		_fdgb := _fccf.ColorAttr
		if _fdgb != nil && _fdgb.ST_HexColorRGB != nil {
			return *_fdgb.ST_HexColorRGB
		}
	}
	return ""
}

// SetAllowOverlapAttr sets the allowOverlap attribute of anchor.
func (_bfe AnchoredDrawing) SetAllowOverlapAttr(val bool) { _bfe._dgc.AllowOverlapAttr = val }

// RStyle returns the name of character style.
// It is defined here http://officeopenxml.com/WPstyleCharStyles.php
func (_ggdde ParagraphProperties) RStyle() string {
	if _ggdde._dfaf.RPr.RStyle != nil {
		return _ggdde._dfaf.RPr.RStyle.ValAttr
	}
	return ""
}

// AddParagraph adds a new paragraph to the document body.
func (_adac *Document) AddParagraph() Paragraph {
	_bcaf := wml.NewEG_BlockLevelElts()
	_adac.doc.Body.EG_BlockLevelElts = append(_adac.doc.Body.EG_BlockLevelElts, _bcaf)
	_eab := wml.NewEG_ContentBlockContent()
	_bcaf.EG_ContentBlockContent = append(_bcaf.EG_ContentBlockContent, _eab)
	_bfb := wml.NewCT_P()
	_eab.P = append(_eab.P, _bfb)
	return Paragraph{_adac, _bfb}
}

// SetOutlineLvl sets outline level of paragraph.
func (_fdcgd Paragraph) SetOutlineLvl(lvl int64) {
	_fdcgd.ensurePPr()
	if _fdcgd._eagd.PPr.OutlineLvl == nil {
		_fdcgd._eagd.PPr.OutlineLvl = wml.NewCT_DecimalNumber()
	}
	_dcbbc := lvl - 1
	_fdcgd._eagd.PPr.OutlineLvl.ValAttr = _dcbbc
}

// FormField is a form within a document. It references the document, so changes
// to the form field wil be reflected in the document if it is saved.
type FormField struct {
	_cbde *wml.CT_FFData
	_gcbd *wml.EG_RunInnerContent
}

// RemoveFootnote removes a footnote from both the paragraph and the document
// the requested footnote must be anchored on the paragraph being referenced.
func (_dccaf Paragraph) RemoveFootnote(id int64) {
	_edfdgb := _dccaf._fagf._beg
	var _gccg int
	for _dfbgb, _ebef := range _edfdgb.CT_Footnotes.Footnote {
		if _ebef.IdAttr == id {
			_gccg = _dfbgb
		}
	}
	_gccg = 0
	_edfdgb.CT_Footnotes.Footnote[_gccg] = nil
	_edfdgb.CT_Footnotes.Footnote[_gccg] = _edfdgb.CT_Footnotes.Footnote[len(_edfdgb.CT_Footnotes.Footnote)-1]
	_edfdgb.CT_Footnotes.Footnote = _edfdgb.CT_Footnotes.Footnote[:len(_edfdgb.CT_Footnotes.Footnote)-1]
	var _gcfbe Run
	for _, _gdee := range _dccaf.Runs() {
		if _bdef, _dagcb := _gdee.IsFootnote(); _bdef {
			if _dagcb == id {
				_gcfbe = _gdee
			}
		}
	}
	_dccaf.RemoveRun(_gcfbe)
}

// ExtractFromFooter returns text from the document footer as an array of TextItems.
func ExtractFromFooter(footer *wml.Ftr) []TextItem { return _dcbb(footer.EG_ContentBlockContent, nil) }

// GetImage returns the ImageRef associated with an AnchoredDrawing.
func (_aeb AnchoredDrawing) GetImage() (common.ImageRef, bool) {
	_cbg := _aeb._dgc.Graphic.GraphicData.Any
	if len(_cbg) > 0 {
		_ea, _acd := _cbg[0].(*picture.Pic)
		if _acd {
			if _ea.BlipFill != nil && _ea.BlipFill.Blip != nil && _ea.BlipFill.Blip.EmbedAttr != nil {
				return _aeb._dg.GetImageByRelID(*_ea.BlipFill.Blip.EmbedAttr)
			}
		}
	}
	return common.ImageRef{}, false
}

// InsertParagraphAfter adds a new empty paragraph after the relativeTo
// paragraph.
func (_agba *Document) InsertParagraphAfter(relativeTo Paragraph) Paragraph {
	return _agba.insertParagraph(relativeTo, false)
}
func (_defdf Paragraph) addSeparateFldChar() *wml.CT_FldChar {
	_fbbc := _defdf.addFldChar()
	_fbbc.FldCharTypeAttr = wml.ST_FldCharTypeSeparate
	return _fbbc
}

// SetPicture sets the watermark picture.
func (_fabae *WatermarkPicture) SetPicture(imageRef common.ImageRef) {
	_ffafa := imageRef.RelID()
	_caca := _fabae.getShape()
	if _fabae._fdgfa != nil {
		_becb := _fabae._fdgfa.EG_ShapeElements
		if len(_becb) > 0 && _becb[0].Imagedata != nil {
			_becb[0].Imagedata.IdAttr = &_ffafa
		}
	} else {
		_bbbc := _fabae.findNode(_caca, "\u0069m\u0061\u0067\u0065\u0064\u0061\u0074a")
		for _fbdea, _cfeb := range _bbbc.Attrs {
			if _cfeb.Name.Local == "\u0069\u0064" {
				_bbbc.Attrs[_fbdea].Value = _ffafa
			}
		}
	}
}

// RunProperties returns the run properties controlling text formatting within the table.
func (_fbcag TableConditionalFormatting) RunProperties() RunProperties {
	if _fbcag._ecbge.RPr == nil {
		_fbcag._ecbge.RPr = wml.NewCT_RPr()
	}
	return RunProperties{_fbcag._ecbge.RPr}
}

// Paragraphs returns the paragraphs defined in an endnote.
func (_cgeg Endnote) Paragraphs() []Paragraph {
	_fdbc := []Paragraph{}
	for _, _dbfc := range _cgeg.content() {
		for _, _abaga := range _dbfc.P {
			_fdbc = append(_fdbc, Paragraph{_cgeg._cceg, _abaga})
		}
	}
	return _fdbc
}
func (_gbgca *Document) putNode(_feac, _cagcd Node, _dcde bool) bool {
	_gbgca.insertImageFromNode(_cagcd)
	_gbgca.insertStyleFromNode(_cagcd)
	switch _acea := _cagcd._ggda.(type) {
	case *Paragraph:
		if _bcfc, _gedd := _feac.X().(*Paragraph); _gedd {
			_gbgca.appendParagraph(_bcfc, *_acea, _dcde)
			return true
		} else {
			for _, _fgbdd := range _feac.Children {
				if _dcff := _gbgca.putNode(_fgbdd, _cagcd, _dcde); _dcff {
					break
				}
			}
		}
	case *Table:
		if _gfgcc, _dgab := _feac.X().(*Paragraph); _dgab {
			_fbba := _gbgca.appendTable(_gfgcc, *_acea, _dcde)
			_fbba.ctTbl = _acea.ctTbl
			return true
		} else {
			for _, _fgaf := range _feac.Children {
				if _bafba := _gbgca.putNode(_fgaf, _cagcd, _dcde); _bafba {
					break
				}
			}
		}
	}
	return false
}

// SetAfter sets the spacing that comes after the paragraph.
func (_fdbeb ParagraphSpacing) SetAfter(after measurement.Distance) {
	_fdbeb._ffede.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
	_fdbeb._ffede.AfterAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(after / measurement.Twips))
}

// AnchorDrawWrapOptions is options to set
// wrapPolygon for wrap text through and tight.
type AnchorDrawWrapOptions struct {
	_cef bool
	_dd  *dml.CT_Point2D
	_cbf []*dml.CT_Point2D
}

// SetTextWrapNone unsets text wrapping so the image can float on top of the
// text. When used in conjunction with X/Y Offset relative to the page it can be
// used to place a logo at the top of a page at an absolute position that
// doesn't interfere with text.
func (_gg AnchoredDrawing) SetTextWrapNone() {
	_gg._dgc.Choice = &wml.WdEG_WrapTypeChoice{}
	_gg._dgc.Choice.WrapNone = wml.NewWdCT_WrapNone()
}

// OnOffValue represents an on/off value that can also be unset
type OnOffValue byte

func (_abaf Styles) initializeStyleDefaults() {
	_dagga := _abaf.AddStyle("\u004e\u006f\u0072\u006d\u0061\u006c", wml.ST_StyleTypeParagraph, true)
	_dagga.SetName("\u004e\u006f\u0072\u006d\u0061\u006c")
	_dagga.SetPrimaryStyle(true)
	_dgfgc := _abaf.AddStyle("D\u0065f\u0061\u0075\u006c\u0074\u0050\u0061\u0072\u0061g\u0072\u0061\u0070\u0068Fo\u006e\u0074", wml.ST_StyleTypeCharacter, true)
	_dgfgc.SetName("\u0044\u0065\u0066\u0061ul\u0074\u0020\u0050\u0061\u0072\u0061\u0067\u0072\u0061\u0070\u0068\u0020\u0046\u006fn\u0074")
	_dgfgc.SetUISortOrder(1)
	_dgfgc.SetSemiHidden(true)
	_dgfgc.SetUnhideWhenUsed(true)
	_gaca := _abaf.AddStyle("\u0054i\u0074\u006c\u0065\u0043\u0068\u0061r", wml.ST_StyleTypeCharacter, false)
	_gaca.SetName("\u0054\u0069\u0074\u006c\u0065\u0020\u0043\u0068\u0061\u0072")
	_gaca.SetBasedOn(_dgfgc.StyleID())
	_gaca.SetLinkedStyle("\u0054\u0069\u0074l\u0065")
	_gaca.SetUISortOrder(10)
	_gaca.RunProperties().Fonts().SetASCIITheme(wml.ST_ThemeMajorAscii)
	_gaca.RunProperties().Fonts().SetEastAsiaTheme(wml.ST_ThemeMajorEastAsia)
	_gaca.RunProperties().Fonts().SetHANSITheme(wml.ST_ThemeMajorHAnsi)
	_gaca.RunProperties().Fonts().SetCSTheme(wml.ST_ThemeMajorBidi)
	_gaca.RunProperties().SetSize(28 * measurement.Point)
	_gaca.RunProperties().SetKerning(14 * measurement.Point)
	_gaca.RunProperties().SetCharacterSpacing(-10 * measurement.Twips)
	_eadbg := _abaf.AddStyle("\u0054\u0069\u0074l\u0065", wml.ST_StyleTypeParagraph, false)
	_eadbg.SetName("\u0054\u0069\u0074l\u0065")
	_eadbg.SetBasedOn(_dagga.StyleID())
	_eadbg.SetNextStyle(_dagga.StyleID())
	_eadbg.SetLinkedStyle(_gaca.StyleID())
	_eadbg.SetUISortOrder(10)
	_eadbg.SetPrimaryStyle(true)
	_eadbg.ParagraphProperties().SetContextualSpacing(true)
	_eadbg.RunProperties().Fonts().SetASCIITheme(wml.ST_ThemeMajorAscii)
	_eadbg.RunProperties().Fonts().SetEastAsiaTheme(wml.ST_ThemeMajorEastAsia)
	_eadbg.RunProperties().Fonts().SetHANSITheme(wml.ST_ThemeMajorHAnsi)
	_eadbg.RunProperties().Fonts().SetCSTheme(wml.ST_ThemeMajorBidi)
	_eadbg.RunProperties().SetSize(28 * measurement.Point)
	_eadbg.RunProperties().SetKerning(14 * measurement.Point)
	_eadbg.RunProperties().SetCharacterSpacing(-10 * measurement.Twips)
	_bcdaa := _abaf.AddStyle("T\u0061\u0062\u006c\u0065\u004e\u006f\u0072\u006d\u0061\u006c", wml.ST_StyleTypeTable, false)
	_bcdaa.SetName("\u004e\u006f\u0072m\u0061\u006c\u0020\u0054\u0061\u0062\u006c\u0065")
	_bcdaa.SetUISortOrder(99)
	_bcdaa.SetSemiHidden(true)
	_bcdaa.SetUnhideWhenUsed(true)
	_bcdaa.X().TblPr = wml.NewCT_TblPrBase()
	_fccg := NewTableWidth()
	_bcdaa.X().TblPr.TblInd = _fccg.X()
	_fccg.SetValue(0 * measurement.Dxa)
	_bcdaa.X().TblPr.TblCellMar = wml.NewCT_TblCellMar()
	_fccg = NewTableWidth()
	_bcdaa.X().TblPr.TblCellMar.Top = _fccg.X()
	_fccg.SetValue(0 * measurement.Dxa)
	_fccg = NewTableWidth()
	_bcdaa.X().TblPr.TblCellMar.Bottom = _fccg.X()
	_fccg.SetValue(0 * measurement.Dxa)
	_fccg = NewTableWidth()
	_bcdaa.X().TblPr.TblCellMar.Left = _fccg.X()
	_fccg.SetValue(108 * measurement.Dxa)
	_fccg = NewTableWidth()
	_bcdaa.X().TblPr.TblCellMar.Right = _fccg.X()
	_fccg.SetValue(108 * measurement.Dxa)
	_eadaa := _abaf.AddStyle("\u004e\u006f\u004c\u0069\u0073\u0074", wml.ST_StyleTypeNumbering, false)
	_eadaa.SetName("\u004eo\u0020\u004c\u0069\u0073\u0074")
	_eadaa.SetUISortOrder(1)
	_eadaa.SetSemiHidden(true)
	_eadaa.SetUnhideWhenUsed(true)
	_deef := []measurement.Distance{16, 13, 12, 11, 11, 11, 11, 11, 11}
	_ebea := []measurement.Distance{240, 40, 40, 40, 40, 40, 40, 40, 40}
	for _cead := 0; _cead < 9; _cead++ {
		_dacf := fmt.Sprintf("\u0048e\u0061\u0064\u0069\u006e\u0067\u0025d", _cead+1)
		_fcbeb := _abaf.AddStyle(_dacf+"\u0043\u0068\u0061\u0072", wml.ST_StyleTypeCharacter, false)
		_fcbeb.SetName(fmt.Sprintf("\u0048e\u0061d\u0069\u006e\u0067\u0020\u0025\u0064\u0020\u0043\u0068\u0061\u0072", _cead+1))
		_fcbeb.SetBasedOn(_dgfgc.StyleID())
		_fcbeb.SetLinkedStyle(_dacf)
		_fcbeb.SetUISortOrder(9 + _cead)
		_fcbeb.RunProperties().SetSize(_deef[_cead] * measurement.Point)
		_eggdg := _abaf.AddStyle(_dacf, wml.ST_StyleTypeParagraph, false)
		_eggdg.SetName(fmt.Sprintf("\u0068\u0065\u0061\u0064\u0069\u006e\u0067\u0020\u0025\u0064", _cead+1))
		_eggdg.SetNextStyle(_dagga.StyleID())
		_eggdg.SetLinkedStyle(_eggdg.StyleID())
		_eggdg.SetUISortOrder(9 + _cead)
		_eggdg.SetPrimaryStyle(true)
		_eggdg.ParagraphProperties().SetKeepNext(true)
		_eggdg.ParagraphProperties().SetSpacing(_ebea[_cead]*measurement.Twips, 0)
		_eggdg.ParagraphProperties().SetOutlineLevel(_cead)
		_eggdg.RunProperties().SetSize(_deef[_cead] * measurement.Point)
	}
}

// Copy makes a deep copy of the document by saving and reading it back.
// It can be useful to avoid sharing common data between two documents.
func (_eead *Document) Copy() (*Document, error) {
	_gabd := bytes.NewBuffer([]byte{})
	_gcg := _eead.save(_gabd, _eead._feg)
	if _gcg != nil {
		return nil, _gcg
	}
	_ggg := _gabd.Bytes()
	_agde := bytes.NewReader(_ggg)
	return _fbee(_agde, int64(_agde.Len()), _eead._feg)
}
func _edbf() *vml.Formulas {
	_bbeaf := vml.NewFormulas()
	_bbeaf.F = []*vml.CT_F{vmldrawing.CreateFormula("\u0069\u0066 \u006c\u0069\u006e\u0065\u0044\u0072\u0061\u0077\u006e\u0020\u0070\u0069\u0078\u0065\u006c\u004c\u0069\u006e\u0065\u0057\u0069\u0064th\u0020\u0030"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0040\u0030\u0020\u0031\u0020\u0030"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0030\u0020\u0030\u0020\u0040\u0031"), vmldrawing.CreateFormula("p\u0072\u006f\u0064\u0020\u0040\u0032\u0020\u0031\u0020\u0032"), vmldrawing.CreateFormula("\u0070r\u006f\u0064\u0020\u0040\u0033\u0020\u0032\u0031\u0036\u0030\u0030 \u0070\u0069\u0078\u0065\u006c\u0057\u0069\u0064\u0074\u0068"), vmldrawing.CreateFormula("\u0070r\u006f\u0064\u0020\u00403\u0020\u0032\u0031\u0036\u00300\u0020p\u0069x\u0065\u006c\u0048\u0065\u0069\u0067\u0068t"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0040\u0030\u0020\u0030\u0020\u0031"), vmldrawing.CreateFormula("p\u0072\u006f\u0064\u0020\u0040\u0036\u0020\u0031\u0020\u0032"), vmldrawing.CreateFormula("\u0070r\u006f\u0064\u0020\u0040\u0037\u0020\u0032\u0031\u0036\u0030\u0030 \u0070\u0069\u0078\u0065\u006c\u0057\u0069\u0064\u0074\u0068"), vmldrawing.CreateFormula("\u0073\u0075\u006d\u0020\u0040\u0038\u0020\u0032\u00316\u0030\u0030\u0020\u0030"), vmldrawing.CreateFormula("\u0070r\u006f\u0064\u0020\u00407\u0020\u0032\u0031\u0036\u00300\u0020p\u0069x\u0065\u006c\u0048\u0065\u0069\u0067\u0068t"), vmldrawing.CreateFormula("\u0073u\u006d \u0040\u0031\u0030\u0020\u0032\u0031\u0036\u0030\u0030\u0020\u0030")}
	return _bbeaf
}

// SetHeadingLevel sets a heading level and style based on the level to a
// paragraph.  The default styles for a new unioffice document support headings
// from level 1 to 8.
func (_egab ParagraphProperties) SetHeadingLevel(idx int) {
	_egab.SetStyle(fmt.Sprintf("\u0048e\u0061\u0064\u0069\u006e\u0067\u0025d", idx))
	if _egab._dfaf.NumPr == nil {
		_egab._dfaf.NumPr = wml.NewCT_NumPr()
	}
	_egab._dfaf.NumPr.Ilvl = wml.NewCT_DecimalNumber()
	_egab._dfaf.NumPr.Ilvl.ValAttr = int64(idx)
}

// SetTopPct sets the cell top margin
func (_gfg CellMargins) SetTopPct(pct float64) {
	_gfg._cdae.Top = wml.NewCT_TblWidth()
	_aff(_gfg._cdae.Top, pct)
}

// ComplexSizeMeasure returns font with its measure which can be mm, cm, in, pt, pc or pi.
func (_efgfe ParagraphProperties) ComplexSizeMeasure() string {
	if _ggeb := _efgfe._dfaf.RPr.SzCs; _ggeb != nil {
		_gebca := _ggeb.ValAttr
		if _gebca.ST_PositiveUniversalMeasure != nil {
			return *_gebca.ST_PositiveUniversalMeasure
		}
	}
	return ""
}

// SetXOffset sets the X offset for an image relative to the origin.
func (_eag AnchoredDrawing) SetXOffset(x measurement.Distance) {
	_eag._dgc.PositionH.Choice = &wml.WdCT_PosHChoice{}
	_eag._dgc.PositionH.Choice.PosOffset = unioffice.Int32(int32(x / measurement.EMU))
}

// ReplaceText replace the text inside node.
func (_gadb *Node) ReplaceText(oldText, newText string) {
	switch _ggedb := _gadb.X().(type) {
	case *Paragraph:
		for _, _egaea := range _ggedb.Runs() {
			for _, _eeeg := range _egaea._adaad.EG_RunInnerContent {
				if _eeeg.T != nil {
					_ffda := _eeeg.T.Content
					_ffda = strings.ReplaceAll(_ffda, oldText, newText)
					_eeeg.T.Content = _ffda
				}
			}
		}
	}
	for _, _daefc := range _gadb.Children {
		_daefc.ReplaceText(oldText, newText)
	}
}

// SetKeepOnOnePage controls if all lines in a paragraph are kept on the same
// page.
func (_ceeg ParagraphProperties) SetKeepOnOnePage(b bool) {
	if !b {
		_ceeg._dfaf.KeepLines = nil
	} else {
		_ceeg._dfaf.KeepLines = wml.NewCT_OnOff()
	}
}

// SetFormat sets the numbering format.
func (_edgd NumberingLevel) SetFormat(f wml.ST_NumberFormat) {
	if _edgd.lvl.NumFmt == nil {
		_edgd.lvl.NumFmt = wml.NewCT_NumFmt()
	}
	_edgd.lvl.NumFmt.ValAttr = f
}
func (_aace Paragraph) addEndBookmark(_dafa int64) *wml.CT_MarkupRange {
	_eccgd := wml.NewEG_PContent()
	_aace._eagd.EG_PContent = append(_aace._eagd.EG_PContent, _eccgd)
	_caccc := wml.NewEG_ContentRunContent()
	_gfce := wml.NewEG_RunLevelElts()
	_febg := wml.NewEG_RangeMarkupElements()
	_cggg := wml.NewCT_MarkupRange()
	_cggg.IdAttr = _dafa
	_febg.BookmarkEnd = _cggg
	_eccgd.EG_ContentRunContent = append(_eccgd.EG_ContentRunContent, _caccc)
	_caccc.EG_RunLevelElts = append(_caccc.EG_RunLevelElts, _gfce)
	_gfce.EG_RangeMarkupElements = append(_gfce.EG_RangeMarkupElements, _febg)
	return _cggg
}

// SetHangingIndent controls the hanging indent of the paragraph.
func (_dddgg ParagraphStyleProperties) SetHangingIndent(m measurement.Distance) {
	if _dddgg._gfee.Ind == nil {
		_dddgg._gfee.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		_dddgg._gfee.Ind.HangingAttr = nil
	} else {
		_dddgg._gfee.Ind.HangingAttr = &sharedTypes.ST_TwipsMeasure{}
		_dddgg._gfee.Ind.HangingAttr.ST_UnsignedDecimalNumber = unioffice.Uint64(uint64(m / measurement.Twips))
	}
}

// SetRight sets the right border to a specified type, color and thickness.
func (_dgbd TableBorders) SetRight(t wml.ST_Border, c color.Color, thickness measurement.Distance) {
	_dgbd._gcdf.Right = wml.NewCT_Border()
	_feadc(_dgbd._gcdf.Right, t, c, thickness)
}
func _ebbad() *vml.Imagedata {
	_ggebb := vml.NewImagedata()
	_gcdbc := "\u0072\u0049\u0064\u0031"
	_ddadff := "\u0057A\u0054\u0045\u0052\u004d\u0041\u0052K"
	_ggebb.IdAttr = &_gcdbc
	_ggebb.TitleAttr = &_ddadff
	return _ggebb
}

// SetCalcOnExit marks if a FormField should be CalcOnExit or not.
func (_dfee FormField) SetCalcOnExit(calcOnExit bool) {
	_cgcd := wml.NewCT_OnOff()
	_cgcd.ValAttr = &sharedTypes.ST_OnOff{Bool: &calcOnExit}
	_dfee._cbde.CalcOnExit = []*wml.CT_OnOff{_cgcd}
}

// SetName marks sets a name attribute for a FormField.
func (_aacbf FormField) SetName(name string) {
	_gecc := wml.NewCT_FFName()
	_gecc.ValAttr = &name
	_aacbf._cbde.Name = []*wml.CT_FFName{_gecc}
}

// SetItalic sets the run to italic.
func (_adcbc RunProperties) SetItalic(b bool) {
	if !b {
		_adcbc._gbdb.I = nil
		_adcbc._gbdb.ICs = nil
	} else {
		_adcbc._gbdb.I = wml.NewCT_OnOff()
		_adcbc._gbdb.ICs = wml.NewCT_OnOff()
	}
}

// SetRight sets the cell right margin
func (_cdb CellMargins) SetRight(d measurement.Distance) {
	_cdb._cdae.Right = wml.NewCT_TblWidth()
	_age(_cdb._cdae.Right, d)
}

// AddBreak adds a line break to a run.
func (_dagg Run) AddBreak() { _bdbb := _dagg.newIC(); _bdbb.Br = wml.NewCT_Br() }
func (_abdd Paragraph) addFldChar() *wml.CT_FldChar {
	_baae := _abdd.AddRun()
	_bfbae := _baae.X()
	_aeac := wml.NewEG_RunInnerContent()
	_bbea := wml.NewCT_FldChar()
	_aeac.FldChar = _bbea
	_bfbae.EG_RunInnerContent = append(_bfbae.EG_RunInnerContent, _aeac)
	return _bbea
}

// SetRightPct sets the cell right margin
func (_gbb CellMargins) SetRightPct(pct float64) {
	_gbb._cdae.Right = wml.NewCT_TblWidth()
	_aff(_gbb._cdae.Right, pct)
}

// X returns the inner wrapped type
func (_cae CellBorders) X() *wml.CT_TcBorders { return _cae._gf }

// X returns the inner wrapped XML type.
func (_eabeg Fonts) X() *wml.CT_Fonts { return _eabeg._feae }

// Font returns the name of paragraph font family.
func (_dgfe ParagraphProperties) Font() string {
	if _daefb := _dgfe._dfaf.RPr.RFonts; _daefb != nil {
		if _daefb.AsciiAttr != nil {
			return *_daefb.AsciiAttr
		} else if _daefb.HAnsiAttr != nil {
			return *_daefb.HAnsiAttr
		} else if _daefb.CsAttr != nil {
			return *_daefb.CsAttr
		}
	}
	return ""
}

// SetToolTip sets the tooltip text for a hyperlink.
func (_ggef HyperLink) SetToolTip(text string) {
	if text == "" {
		_ggef._baaf.TooltipAttr = nil
	} else {
		_ggef._baaf.TooltipAttr = unioffice.String(text)
	}
}

// X returns the inner wrapped XML type.
func (_abb Cell) X() *wml.CT_Tc { return _abb._gge }

// SetStyle sets the style of a paragraph.
func (_gbefd ParagraphProperties) SetStyle(s string) {
	if s == "" {
		_gbefd._dfaf.PStyle = nil
	} else {
		_gbefd._dfaf.PStyle = wml.NewCT_String()
		_gbefd._dfaf.PStyle.ValAttr = s
	}
}
func (_gfed *Document) appendTable(_bafc *Paragraph, _adab Table, _bcca bool) Table {
	_gdff := _gfed.doc.Body
	_gfdc := wml.NewEG_BlockLevelElts()
	_gfed.doc.Body.EG_BlockLevelElts = append(_gfed.doc.Body.EG_BlockLevelElts, _gfdc)
	_adee := wml.NewEG_ContentBlockContent()
	_gfdc.EG_ContentBlockContent = append(_gfdc.EG_ContentBlockContent, _adee)
	if _bafc != nil {
		_efed := _bafc.X()
		for _ddc, _afc := range _gdff.EG_BlockLevelElts {
			for _, _fde := range _afc.EG_ContentBlockContent {
				for _cgd, _eddg := range _adee.P {
					if _eddg == _efed {
						_aef := _adab.X()
						_ggaf := wml.NewEG_BlockLevelElts()
						_fdbe := wml.NewEG_ContentBlockContent()
						_ggaf.EG_ContentBlockContent = append(_ggaf.EG_ContentBlockContent, _fdbe)
						_fdbe.Tbl = append(_fdbe.Tbl, _aef)
						_gdff.EG_BlockLevelElts = append(_gdff.EG_BlockLevelElts, nil)
						if _bcca {
							copy(_gdff.EG_BlockLevelElts[_ddc+1:], _gdff.EG_BlockLevelElts[_ddc:])
							_gdff.EG_BlockLevelElts[_ddc] = _ggaf
							if _cgd != 0 {
								_egdb := wml.NewEG_BlockLevelElts()
								_eadb := wml.NewEG_ContentBlockContent()
								_egdb.EG_ContentBlockContent = append(_egdb.EG_ContentBlockContent, _eadb)
								_eadb.P = _fde.P[:_cgd]
								_gdff.EG_BlockLevelElts = append(_gdff.EG_BlockLevelElts, nil)
								copy(_gdff.EG_BlockLevelElts[_ddc+1:], _gdff.EG_BlockLevelElts[_ddc:])
								_gdff.EG_BlockLevelElts[_ddc] = _egdb
							}
							_fde.P = _fde.P[_cgd:]
						} else {
							copy(_gdff.EG_BlockLevelElts[_ddc+2:], _gdff.EG_BlockLevelElts[_ddc+1:])
							_gdff.EG_BlockLevelElts[_ddc+1] = _ggaf
							if _cgd != len(_fde.P)-1 {
								_ebeb := wml.NewEG_BlockLevelElts()
								_egea := wml.NewEG_ContentBlockContent()
								_ebeb.EG_ContentBlockContent = append(_ebeb.EG_ContentBlockContent, _egea)
								_egea.P = _fde.P[_cgd+1:]
								_gdff.EG_BlockLevelElts = append(_gdff.EG_BlockLevelElts, nil)
								copy(_gdff.EG_BlockLevelElts[_ddc+3:], _gdff.EG_BlockLevelElts[_ddc+2:])
								_gdff.EG_BlockLevelElts[_ddc+2] = _ebeb
							}
							_fde.P = _fde.P[:_cgd+1]
						}
						break
					}
				}
				for _, _gdfg := range _fde.Tbl {
					_dadc := _adaa(_gdfg, _efed, _bcca)
					if _dadc != nil {
						break
					}
				}
			}
		}
	} else {
		_adee.Tbl = append(_adee.Tbl, _adab.X())
	}
	return Table{_gfed, _adab.X()}
}

// FormFields extracts all of the fields from a document.  They can then be
// manipulated via the methods on the field and the document saved.
func (_abg *Document) FormFields() []FormField {
	_dccf := []FormField{}
	for _, _baea := range _abg.Paragraphs() {
		_cfg := _baea.Runs()
		for _bdga, _ecgd := range _cfg {
			for _, _ffbgb := range _ecgd._adaad.EG_RunInnerContent {
				if _ffbgb.FldChar == nil || _ffbgb.FldChar.FfData == nil {
					continue
				}
				if _ffbgb.FldChar.FldCharTypeAttr == wml.ST_FldCharTypeBegin {
					if len(_ffbgb.FldChar.FfData.Name) == 0 || _ffbgb.FldChar.FfData.Name[0].ValAttr == nil {
						continue
					}
					_ggdd := FormField{_cbde: _ffbgb.FldChar.FfData}
					if _ffbgb.FldChar.FfData.TextInput != nil {
						for _dbfff := _bdga + 1; _dbfff < len(_cfg)-1; _dbfff++ {
							if len(_cfg[_dbfff]._adaad.EG_RunInnerContent) == 0 {
								continue
							}
							_cdaa := _cfg[_dbfff]._adaad.EG_RunInnerContent[0]
							if _cdaa.FldChar != nil && _cdaa.FldChar.FldCharTypeAttr == wml.ST_FldCharTypeSeparate {
								if len(_cfg[_dbfff+1]._adaad.EG_RunInnerContent) == 0 {
									continue
								}
								if _cfg[_dbfff+1]._adaad.EG_RunInnerContent[0].FldChar == nil {
									_ggdd._gcbd = _cfg[_dbfff+1]._adaad.EG_RunInnerContent[0]
									break
								}
							}
						}
					}
					_dccf = append(_dccf, _ggdd)
				}
			}
		}
	}
	return _dccf
}
