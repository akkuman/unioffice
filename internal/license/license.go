package license

import (
	"bytes"
	"compress/gzip"
	"crypto"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"crypto/rsa"
	"crypto/sha256"
	"crypto/sha512"
	"crypto/x509"
	"encoding/base64"
	"encoding/binary"
	"encoding/hex"
	"encoding/json"
	"encoding/pem"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/unidoc/unioffice/common"
	"github.com/unidoc/unioffice/common/logger"
)

const UNIOFFICE_CUSTOMER_NAME = "UNIOFFICE_CUSTOMER_NAME"

type defaultStateHolder struct{}

func GetLicenseKey() *LicenseKey {
	if unLicKey == nil {
		return nil
	}
	_cee := *unLicKey
	return &_cee
}

type meteredUsageCheckinForm struct {
	Instance       string         `json:"inst"`
	Next           string         `json:"next"`
	UsageNumber    int            `json:"usage_number"`
	NumFailed      int64          `json:"num_failed"`
	Hostname       string         `json:"hostname"`
	LocalIP        string         `json:"local_ip"`
	MacAddress     string         `json:"mac_address"`
	Package        string         `json:"package"`
	PackageVersion string         `json:"package_version"`
	Usage          map[string]int `json:"u"`
}

func _cge(_agb, _eda []byte) ([]byte, error) {
	_dfa, _gae := aes.NewCipher(_agb)
	if _gae != nil {
		return nil, _gae
	}
	_bdc := make([]byte, aes.BlockSize+len(_eda))
	_bdb := _bdc[:aes.BlockSize]
	if _, _aeb := io.ReadFull(rand.Reader, _bdb); _aeb != nil {
		return nil, _aeb
	}
	_fgde := cipher.NewCFBEncrypter(_dfa, _bdb)
	_fgde.XORKeyStream(_bdc[aes.BlockSize:], _eda)
	_eebe := make([]byte, base64.URLEncoding.EncodedLen(len(_bdc)))
	base64.URLEncoding.Encode(_eebe, _bdc)
	return _eebe, nil
}

var _ege stateLoader = defaultStateHolder{}

type LicenseKey struct {
	LicenseId         string    `json:"license_id"`
	CustomerId        string    `json:"customer_id"`
	CustomerName      string    `json:"customer_name"`
	Tier              string    `json:"tier"`
	CreatedAt         time.Time `json:"-"`
	CreatedAtInt      int64     `json:"created_at"`
	ExpiresAt         time.Time `json:"-"`
	ExpiresAtInt      int64     `json:"expires_at"`
	CreatedBy         string    `json:"created_by"`
	CreatorName       string    `json:"creator_name"`
	CreatorEmail      string    `json:"creator_email"`
	UniPDF            bool      `json:"unipdf"`
	UniOffice         bool      `json:"unioffice"`
	UniHTML           bool      `json:"unihtml"`
	Trial             bool      `json:"trial"`
	unknownBoolField1 bool
	unknownStrField2  string
}
type reportState struct {
	Instance      string         `json:"inst"`
	Next          string         `json:"n"`
	Docs          int64          `json:"d"`
	NumErrors     int64          `json:"e"`
	LimitDocs     bool           `json:"ld"`
	RemainingDocs int64          `json:"rd"`
	LastReported  time.Time      `json:"lr"`
	LastWritten   time.Time      `json:"lw"`
	Usage         map[string]int `json:"u"`
}

func GetHwaddrAndNetips() ([]string, []string, error) {
	netIfaces, err := net.Interfaces()
	if err != nil {
		return nil, nil, err
	}
	var hwAddrs []string
	var netips []string
	for _, netIface := range netIfaces {
		if netIface.Flags&net.FlagUp == 0 || bytes.Equal(netIface.HardwareAddr, nil) {
			continue
		}
		addrs, err := netIface.Addrs()
		if err != nil {
			return nil, nil, err
		}
		i := 0
		for _, addr := range addrs {
			var netIP net.IP
			switch v := addr.(type) {
			case *net.IPNet:
				netIP = v.IP
			case *net.IPAddr:
				netIP = v.IP
			}
			if netIP.IsLoopback() {
				continue
			}
			if netIP.To4() == nil {
				continue
			}
			netips = append(netips, netIP.String())
			i++
		}
		hwAddr := netIface.HardwareAddr.String()
		if hwAddr != "" && i > 0 {
			hwAddrs = append(hwAddrs, hwAddr)
		}
	}
	return hwAddrs, netips, nil
}
func _ae(_feg string, _aef []byte) (string, error) {
	_dga, _ := pem.Decode([]byte(_feg))
	if _dga == nil {
		return "", fmt.Errorf("\u0050\u0072\u0069\u0076\u004b\u0065\u0079\u0020\u0066a\u0069\u006c\u0065\u0064")
	}
	_agcg, _fea := x509.ParsePKCS1PrivateKey(_dga.Bytes)
	if _fea != nil {
		return "", _fea
	}
	_be := sha512.New()
	_be.Write(_aef)
	_bb := _be.Sum(nil)
	_faf, _fea := rsa.SignPKCS1v15(rand.Reader, _agcg, crypto.SHA512, _bb)
	if _fea != nil {
		return "", _fea
	}
	_efb := base64.StdEncoding.EncodeToString(_aef)
	_efb += "\u000a\u002b\u000a"
	_efb += base64.StdEncoding.EncodeToString(_faf)
	return _efb, nil
}

const (
	_ff  = "\u002d\u002d\u002d--\u0042\u0045\u0047\u0049\u004e\u0020\u0055\u004e\u0049D\u004fC\u0020L\u0049C\u0045\u004e\u0053\u0045\u0020\u004b\u0045\u0059\u002d\u002d\u002d\u002d\u002d"
	_gga = "\u002d\u002d\u002d\u002d\u002d\u0045\u004e\u0044\u0020\u0055\u004e\u0049\u0044\u004f\u0043 \u004cI\u0043\u0045\u004e\u0053\u0045\u0020\u004b\u0045\u0059\u002d\u002d\u002d\u002d\u002d"
)

type meteredStatusForm struct{}

const _gge = "\u0033\u0030\u0035\u0063\u0033\u0030\u0030\u00640\u0036\u0030\u0039\u0032\u0061\u0038\u00364\u0038\u0038\u0036\u0066\u0037\u0030d\u0030\u0031\u0030\u0031\u0030\u00310\u0035\u0030\u0030\u0030\u0033\u0034\u0062\u0030\u0030\u0033\u0030\u00348\u0030\u0032\u0034\u0031\u0030\u0030\u0062\u0038\u0037\u0065\u0061\u0066\u0062\u0036\u0063\u0030\u0037\u0034\u0039\u0039\u0065\u0062\u00397\u0063\u0063\u0039\u0064\u0033\u0035\u0036\u0035\u0065\u0063\u00663\u0031\u0036\u0038\u0031\u0039\u0036\u0033\u0030\u0031\u0039\u0030\u0037c\u0038\u0034\u0031\u0061\u0064\u0064c6\u0036\u0035\u0030\u0038\u0036\u0062\u0062\u0033\u0065\u0064\u0038\u0065\u0062\u0031\u0032\u0064\u0039\u0064\u0061\u0032\u0036\u0063\u0061\u0066\u0061\u0039\u0036\u00345\u0030\u00314\u0036\u0064\u0061\u0038\u0062\u0064\u0030\u0063c\u0066\u0031\u0035\u0035\u0066\u0063a\u0063\u0063\u00368\u0036\u0039\u0035\u0035\u0065\u0066\u0030\u0033\u0030\u0032\u0066\u0061\u0034\u0034\u0061\u0061\u0033\u0065\u0063\u0038\u0039\u0034\u0031\u0037\u0062\u0030\u0032\u0030\u0033\u0030\u0031\u0030\u0030\u0030\u0031"

var _edf = false

func _gbc(resp *http.Response) ([]byte, error) {
	var res []byte
	reader, err := _cebe(resp)
	if err != nil {
		return res, err
	}
	return ioutil.ReadAll(reader)
}
func SetLegacyLicenseKey(s string) error {
	_gff := regexp.MustCompile("\u005c\u0073")
	s = _gff.ReplaceAllString(s, "")
	var _daa io.Reader
	_daa = strings.NewReader(s)
	_daa = base64.NewDecoder(base64.RawURLEncoding, _daa)
	_daa, _ebe := gzip.NewReader(_daa)
	if _ebe != nil {
		return _ebe
	}
	_fee := json.NewDecoder(_daa)
	_faga := &LegacyLicense{}
	if _gcd := _fee.Decode(_faga); _gcd != nil {
		return _gcd
	}
	if _bbc := _faga.Verify(_aa); _bbc != nil {
		return errors.New("\u006c\u0069\u0063en\u0073\u0065\u0020\u0076\u0061\u006c\u0069\u0064\u0061\u0074\u0069\u006e\u0020\u0065\u0072\u0072\u006f\u0072")
	}
	if _faga.Expiration.Before(common.ReleasedAt) {
		return errors.New("\u006ci\u0063e\u006e\u0073\u0065\u0020\u0065\u0078\u0070\u0069\u0072\u0065\u0064")
	}
	_gee := time.Now().UTC()
	_abeg := LicenseKey{}
	_abeg.CreatedAt = _gee
	_abeg.CustomerId = "\u004c\u0065\u0067\u0061\u0063\u0079"
	_abeg.CustomerName = _faga.Name
	_abeg.Tier = LicenseTierBusiness
	_abeg.ExpiresAt = _faga.Expiration
	_abeg.CreatorName = "\u0055\u006e\u0069\u0044\u006f\u0063\u0020\u0073\u0075p\u0070\u006f\u0072\u0074"
	_abeg.CreatorEmail = "\u0073\u0075\u0070\u0070\u006f\u0072\u0074\u0040\u0075\u006e\u0069\u0064o\u0063\u002e\u0069\u006f"
	_abeg.UniOffice = true
	unLicKey = &_abeg
	return nil
}
func _cebe(_acab *http.Response) (io.ReadCloser, error) {
	var _edcd error
	var _egdf io.ReadCloser
	switch strings.ToLower(_acab.Header.Get("\u0043\u006fn\u0074\u0065\u006et\u002d\u0045\u006e\u0063\u006f\u0064\u0069\u006e\u0067")) {
	case "\u0067\u007a\u0069\u0070":
		_egdf, _edcd = gzip.NewReader(_acab.Body)
		if _edcd != nil {
			return _egdf, _edcd
		}
		defer _egdf.Close()
	default:
		_egdf = _acab.Body
	}
	return _egdf, nil
}
func SetLicenseKey(content string, customerName string) error {
	if _edf {
		return nil
	}
	_fddd, _fefa := _dedb(content)
	if _fefa != nil {
		logger.Log.Error("\u004c\u0069c\u0065\u006e\u0073\u0065\u0020\u0063\u006f\u0064\u0065\u0020\u0064\u0065\u0063\u006f\u0064\u0065\u0020\u0065\u0072\u0072\u006f\u0072: \u0025\u0076", _fefa)
		return _fefa
	}
	if !strings.EqualFold(_fddd.CustomerName, customerName) {
		logger.Log.Error("L\u0069ce\u006es\u0065 \u0063\u006f\u0064\u0065\u0020i\u0073\u0073\u0075e\u0020\u002d\u0020\u0043\u0075s\u0074\u006f\u006de\u0072\u0020\u006e\u0061\u006d\u0065\u0020\u006d\u0069\u0073\u006da\u0074\u0063\u0068, e\u0078\u0070\u0065\u0063\u0074\u0065d\u0020\u0027\u0025\u0073\u0027\u002c\u0020\u0062\u0075\u0074\u0020\u0067o\u0074 \u0027\u0025\u0073\u0027", customerName, _fddd.CustomerName)
		return fmt.Errorf("\u0063\u0075\u0073\u0074\u006fm\u0065\u0072\u0020\u006e\u0061\u006d\u0065\u0020\u006d\u0069\u0073\u006d\u0061t\u0063\u0068\u002c\u0020\u0065\u0078\u0070\u0065\u0063\u0074\u0065\u0064\u0020\u0027\u0025\u0073\u0027\u002c\u0020\u0062\u0075\u0074\u0020\u0067\u006f\u0074\u0020\u0027\u0025\u0073'", customerName, _fddd.CustomerName)
	}
	_fefa = _fddd.Validate()
	if _fefa != nil {
		logger.Log.Error("\u004c\u0069\u0063\u0065\u006e\u0073e\u0020\u0063\u006f\u0064\u0065\u0020\u0076\u0061\u006c\u0069\u0064\u0061\u0074i\u006f\u006e\u0020\u0065\u0072\u0072\u006fr\u003a\u0020\u0025\u0076", _fefa)
		return _fefa
	}
	unLicKey = &_fddd
	return nil
}
func Track(docKey string, useKey string) error { return track(docKey, useKey, false) }
func (_age *LicenseKey) isExpired() bool       { return _age.getExpiryDateToCompare().After(_age.ExpiresAt) }
func init() {
	_fca, _ea := hex.DecodeString(_gge)
	if _ea != nil {
		log.Fatalf("e\u0072\u0072\u006f\u0072 r\u0065a\u0064\u0069\u006e\u0067\u0020k\u0065\u0079\u003a\u0020\u0025\u0073", _ea)
	}
	_dfg, _ea := x509.ParsePKIXPublicKey(_fca)
	if _ea != nil {
		log.Fatalf("e\u0072\u0072\u006f\u0072 r\u0065a\u0064\u0069\u006e\u0067\u0020k\u0065\u0079\u003a\u0020\u0025\u0073", _ea)
	}
	_aa = _dfg.(*rsa.PublicKey)
}

type meteredStatusResp struct {
	Valid        bool  `json:"valid"`
	OrgCredits   int64 `json:"org_credits"`
	OrgUsed      int64 `json:"org_used"`
	OrgRemaining int64 `json:"org_remaining"`
}

var gMap2 map[string]struct{}
var time20100101 = time.Date(2010, 1, 1, 0, 0, 0, 0, time.UTC)

func GenRefId(prefix string) (string, error) {
	var buf bytes.Buffer
	buf.WriteString(prefix)
	data := make([]byte, 8+16)
	timestamp := time.Now().UTC().UnixNano()
	binary.BigEndian.PutUint64(data, uint64(timestamp))
	_, err := rand.Read(data[8:])
	if err != nil {
		return "", err
	}
	buf.WriteString(hex.EncodeToString(data))
	return buf.String(), nil
}

const UNIOFFICE_LICENSE_PATH = "UNIOFFICE_LICENSE_PATH"

func SetMeteredKey(apiKey string) error {
	if len(apiKey) == 0 {
		logger.Log.Error("\u004d\u0065\u0074\u0065\u0072e\u0064\u0020\u004c\u0069\u0063\u0065\u006e\u0073\u0065\u0020\u0041\u0050\u0049 \u004b\u0065\u0079\u0020\u006d\u0075\u0073\u0074\u0020\u006e\u006f\u0074\u0020\u0062\u0065\u0020\u0065\u006d\u0070\u0074\u0079")
		logger.Log.Error("\u002d\u0020\u0047\u0072\u0061\u0062\u0020\u006f\u006e\u0065\u0020\u0069\u006e\u0020\u0074h\u0065\u0020\u0046\u0072\u0065\u0065\u0020\u0054\u0069\u0065\u0072\u0020\u0061t\u0020\u0068\u0074\u0074\u0070\u0073\u003a\u002f\u002f\u0063\u006c\u006fud\u002e\u0075\u006e\u0069\u0064\u006f\u0063\u002e\u0069\u006f")
		return fmt.Errorf("\u006de\u0074\u0065\u0072e\u0064\u0020\u006ci\u0063en\u0073\u0065\u0020\u0061\u0070\u0069\u0020k\u0065\u0079\u0020\u006d\u0075\u0073\u0074\u0020\u006e\u006f\u0074\u0020\u0062\u0065\u0020\u0065\u006d\u0070\u0074\u0079\u003a\u0020\u0063\u0072\u0065\u0061\u0074\u0065 o\u006ee\u0020\u0061\u0074\u0020\u0068\u0074t\u0070\u0073\u003a\u002f\u002fc\u006c\u006f\u0075\u0064\u002e\u0075\u006e\u0069\u0064\u006f\u0063.\u0069\u006f")
	}
	if unLicKey != nil && (unLicKey.unknownBoolField1 || unLicKey.Tier != LicenseTierUnlicensed) {
		logger.Log.Error("\u0045\u0052\u0052\u004f\u0052:\u0020\u0043\u0061\u006e\u006eo\u0074 \u0073\u0065\u0074\u0020\u006c\u0069\u0063\u0065\u006e\u0073\u0065\u0020\u006b\u0065\u0079\u0020\u0074\u0077\u0069c\u0065\u0020\u002d\u0020\u0053\u0068\u006f\u0075\u006c\u0064\u0020\u006a\u0075\u0073\u0074\u0020\u0069\u006e\u0069\u0074\u0069\u0061\u006c\u0069z\u0065\u0020\u006f\u006e\u0063\u0065")
		return errors.New("\u006c\u0069\u0063en\u0073\u0065\u0020\u006b\u0065\u0079\u0020\u0061\u006c\u0072\u0065\u0061\u0064\u0079\u0020\u0073\u0065\u0074")
	}
	_dfd := _cgf()
	_dfd._dac = apiKey
	_bae, _edd := _dfd.getStatus()
	if _edd != nil {
		return _edd
	}
	if !_bae.Valid {
		return errors.New("\u006b\u0065\u0079\u0020\u006e\u006f\u0074\u0020\u0076\u0061\u006c\u0069\u0064")
	}
	_agef := &LicenseKey{unknownBoolField1: true, unknownStrField2: apiKey}
	unLicKey = _agef
	return nil
}

const publicKey = `
-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAmFUiyd7b5XjpkP5Rap4w
Dc1dyzIQ4LekxrvytnEMpNUbo6iA74V8ruZOvrScsf2QeN9/qrUG8qEbUWdoEYq+
otFNAFNxlGbxbDHcdGVaM0OXdXgDyL5aIEagL0c5pwjIdPGIn46f78eMJ+JkdcpD
DJaqYXdrz5KeshjSiIaa7menBIAXS4UFxNfHhN0HCYZYqQG7bK+s5rRHonydNWEG
H8Myvr2pya2KrMumfmAxUB6fenC/4O0Wr8gfPOU8RitmbDvQPIRXOL4vTBrBdbaA
9nwNP+i//20MT2bxmeWB+gpcEhGpXZ733azQxrC3J4v3CZmENStDK/KDSPKUGfu6
fwIDAQAB
-----END PUBLIC KEY-----
`

type stateLoader interface {
	loadState(_fede string) (reportState, error)
	updateState(_cbf, _efef, _aegd string, _cccg int, _ebcg bool, _ecdg int, _dea int, _aaa time.Time, _ddc map[string]int) error
}

var unLicKey = MakeUnlicensedKey()

type LegacyLicenseType byte

func _eae(key, _dba []byte) ([]byte, error) {
	_bcd := make([]byte, base64.URLEncoding.DecodedLen(len(_dba)))
	n, err := base64.URLEncoding.Decode(_bcd, _dba)
	if err != nil {
		return nil, err
	}
	_bcd = _bcd[:n]
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	if len(_bcd) < aes.BlockSize {
		return nil, errors.New("ciphertext too short")
	}
	iv := _bcd[:aes.BlockSize]
	_bcd = _bcd[aes.BlockSize:]
	cipherStream := cipher.NewCFBDecrypter(block, iv)
	cipherStream.XORKeyStream(_bcd, _bcd)
	return _bcd, nil
}

var mutex = &sync.Mutex{}

type LegacyLicense struct {
	Name        string
	Signature   string `json:",omitempty"`
	Expiration  time.Time
	LicenseType LegacyLicenseType
}

func gzipData(body []byte) (io.Reader, error) {
	buf := new(bytes.Buffer)
	w := gzip.NewWriter(buf)
	w.Write(body)
	err := w.Close()
	if err != nil {
		return nil, err
	}
	return buf, nil
}

func (lk *LicenseKey) Validate() error {
	if lk.unknownBoolField1 {
		return nil
	}
	if len(lk.LicenseId) < 10 {
		return fmt.Errorf("invalid license: License Id")
	}
	if len(lk.CustomerId) < 10 {
		return fmt.Errorf("invalid license: Customer Id")
	}
	if len(lk.CustomerName) < 1 {
		return fmt.Errorf("invalid license: Customer Name")
	}
	if time20100101.After(lk.CreatedAt) {
		return fmt.Errorf("invalid license: Created At is invalid")
	}
	if lk.ExpiresAt.IsZero() {
		_ad := lk.CreatedAt.AddDate(1, 0, 0)
		if _efg.After(_ad) {
			_ad = _efg
		}
		lk.ExpiresAt = _ad
	}
	if lk.CreatedAt.After(lk.ExpiresAt) {
		return fmt.Errorf("invalid license: Created At cannot be Greater than Expires At")
	}
	if lk.isExpired() {
		return fmt.Errorf("invalid license: The license has already expired")
	}
	if len(lk.CreatorName) < 1 {
		return fmt.Errorf("invalid license: Creator name")
	}
	if len(lk.CreatorEmail) < 1 {
		return fmt.Errorf("invalid license: Creator email")
	}
	if lk.CreatedAt.After(time20190606) {
		if !lk.UniOffice {
			return fmt.Errorf("invalid license: This UniDoc key is invalid for UniOffice")
		}
	}
	return nil
}
func track(docKey string, useKey string, boolv bool) error {
	if unLicKey == nil {
		return errors.New("\u006e\u006f\u0020\u006c\u0069\u0063\u0065\u006e\u0073e\u0020\u006b\u0065\u0079")
	}
	if !unLicKey.unknownBoolField1 || len(unLicKey.unknownStrField2) == 0 {
		return nil
	}
	if len(docKey) == 0 && !boolv {
		return errors.New("\u0064\u006f\u0063\u004b\u0065\u0079\u0020\u006e\u006ft\u0020\u0073\u0065\u0074")
	}
	mutex.Lock()
	defer mutex.Unlock()
	if gMap2 == nil {
		gMap2 = map[string]struct{}{}
	}
	if gMap1 == nil {
		gMap1 = map[string]int{}
	}
	i := 0
	if !boolv {
		_, ok := gMap2[docKey]
		if !ok {
			gMap2[docKey] = struct{}{}
			i++
		}
		if i == 0 {
			return nil
		}
		gMap1[useKey]++
	}
	nowTime := time.Now()
	_reportState, err := _ege.loadState(unLicKey.unknownStrField2)
	if err != nil {
		logger.Log.Error("ERROR: %v", err)
		return err
	}
	if _reportState.Usage == nil {
		_reportState.Usage = map[string]int{}
	}
	for k, v := range gMap1 {
		_reportState.Usage[k] += v
	}
	gMap1 = nil
	const hour24 = 24 * time.Hour
	const day3 = 3 * 24 * time.Hour
	if len(_reportState.Instance) == 0 || nowTime.Sub(_reportState.LastReported) > hour24 || (_reportState.LimitDocs && _reportState.RemainingDocs <= _reportState.Docs+int64(i)) || boolv {
		hostname, err := os.Hostname()
		if err != nil {
			return err
		}
		docs := _reportState.Docs
		hwAddrs, netIPs, err := GetHwaddrAndNetips()
		if err != nil {
			return err
		}
		sort.Strings(netIPs)
		sort.Strings(hwAddrs)
		_egegf, err := _afce()
		if err != nil {
			return err
		}
		_eab := false
		for _, netIP := range netIPs {
			if netIP == _egegf.String() {
				_eab = true
			}
		}
		if !_eab {
			netIPs = append(netIPs, _egegf.String())
		}
		_afbb := _cgf()
		_afbb._dac = unLicKey.unknownStrField2
		docs += int64(i)
		_dfdc := meteredUsageCheckinForm{Instance: _reportState.Instance, Next: _reportState.Next, UsageNumber: int(docs), NumFailed: _reportState.NumErrors, Hostname: hostname, LocalIP: strings.Join(netIPs, "\u002c\u0020"), MacAddress: strings.Join(hwAddrs, "\u002c\u0020"), Package: "\u0075n\u0069\u006f\u0066\u0066\u0069\u0063e", PackageVersion: common.Version, Usage: _reportState.Usage}
		if len(hwAddrs) == 0 {
			_dfdc.MacAddress = "\u006e\u006f\u006e\u0065"
		}
		_bfaf := int64(0)
		_bbd := _reportState.NumErrors
		_ffb := nowTime
		_geg := 0
		_ged := _reportState.LimitDocs
		_ecb, err := _afbb.checkinUsage(_dfdc)
		if err != nil {
			if nowTime.Sub(_reportState.LastReported) > day3 {
				return errors.New("\u0074\u006f\u006f\u0020\u006c\u006f\u006e\u0067\u0020\u0073\u0069\u006e\u0063\u0065\u0020\u006c\u0061\u0073\u0074\u0020\u0073\u0075\u0063\u0063e\u0073\u0073\u0066\u0075\u006c \u0063\u0068e\u0063\u006b\u0069\u006e")
			}
			_bfaf = docs
			_bbd++
			_ffb = _reportState.LastReported
		} else {
			_ged = _ecb.LimitDocs
			_geg = _ecb.RemainingDocs
			_bbd = 0
		}
		if len(_ecb.Instance) == 0 {
			_ecb.Instance = _dfdc.Instance
		}
		if len(_ecb.Next) == 0 {
			_ecb.Next = _dfdc.Next
		}
		err = _ege.updateState(_afbb._dac, _ecb.Instance, _ecb.Next, int(_bfaf), _ged, _geg, int(_bbd), _ffb, nil)
		if err != nil {
			return err
		}
		if !_ecb.Success {
			return fmt.Errorf("\u0065r\u0072\u006f\u0072\u003a\u0020\u0025s", _ecb.Message)
		}
	} else {
		err = _ege.updateState(unLicKey.unknownStrField2, _reportState.Instance, _reportState.Next, int(_reportState.Docs)+i, _reportState.LimitDocs, int(_reportState.RemainingDocs), int(_reportState.NumErrors), _reportState.LastReported, _reportState.Usage)
		if err != nil {
			return err
		}
	}
	return nil
}
func (mc *meteredClient) getStatus() (meteredStatusResp, error) {
	var resp meteredStatusResp
	u := mc.baseURL + "/metered/status"
	var form meteredStatusForm
	_ecd, err := json.Marshal(form)
	if err != nil {
		return resp, err
	}
	bodyReader, err := gzipData(_ecd)
	if err != nil {
		return resp, err
	}
	req, err := http.NewRequest("POST", u, bodyReader)
	if err != nil {
		return resp, err
	}
	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("Content-Encoding", "gzip")
	req.Header.Add("\u0041c\u0063e\u0070\u0074\u002d\u0045\u006e\u0063\u006f\u0064\u0069\u006e\u0067", "\u0067\u007a\u0069\u0070")
	req.Header.Add("\u0058-\u0041\u0050\u0049\u002d\u004b\u0045Y", mc._dac)
	pResponse, err := mc.httpClient.Do(req)
	if err != nil {
		return resp, err
	}
	defer pResponse.Body.Close()
	if pResponse.StatusCode != 200 {
		return resp, fmt.Errorf("failed to checkin, status code is: %d", pResponse.StatusCode)
	}
	_ac, err := _gbc(pResponse)
	if err != nil {
		return resp, err
	}
	err = json.Unmarshal(_ac, &resp)
	if err != nil {
		return resp, err
	}
	return resp, nil
}

var time20190606 = time.Date(2019, 6, 6, 0, 0, 0, 0, time.UTC)

func (_fg LegacyLicense) Verify(pubKey *rsa.PublicKey) error {
	_fb := _fg
	_fb.Signature = ""
	buf := bytes.Buffer{}
	_bed := json.NewEncoder(&buf)
	if _bbb := _bed.Encode(_fb); _bbb != nil {
		return _bbb
	}
	_eee, _cfg := hex.DecodeString(_fg.Signature)
	if _cfg != nil {
		return _cfg
	}
	_ceb := sha256.Sum256(buf.Bytes())
	_cfg = rsa.VerifyPKCS1v15(pubKey, crypto.SHA256, _ceb[:], _eee)
	return _cfg
}
func _dedb(_gf string) (LicenseKey, error) {
	var lk LicenseKey
	_fcd, _dbf := _bad(_ff, _gga, _gf)
	if _dbf != nil {
		return lk, _dbf
	}
	_ebc, _dbf := _efe(publicKey, _fcd)
	if _dbf != nil {
		return lk, _dbf
	}
	_dbf = json.Unmarshal(_ebc, &lk)
	if _dbf != nil {
		return lk, _dbf
	}
	lk.CreatedAt = time.Unix(lk.CreatedAtInt, 0)
	if lk.ExpiresAtInt > 0 {
		_edb := time.Unix(lk.ExpiresAtInt, 0)
		lk.ExpiresAt = _edb
	}
	return lk, nil
}
func TrackUse(useKey string) {
	if unLicKey == nil {
		return
	}
	if !unLicKey.unknownBoolField1 || len(unLicKey.unknownStrField2) == 0 {
		return
	}
	if len(useKey) == 0 {
		return
	}
	mutex.Lock()
	defer mutex.Unlock()
	if gMap1 == nil {
		gMap1 = map[string]int{}
	}
	gMap1[useKey]++
}

type meteredClient struct {
	baseURL    string
	_dac       string
	httpClient *http.Client
}
type meteredUsageCheckinResp struct {
	Instance      string `json:"inst"`
	Next          string `json:"next"`
	Success       bool   `json:"success"`
	Message       string `json:"message"`
	RemainingDocs int    `json:"rd"`
	LimitDocs     bool   `json:"ld"`
}

func GetMeteredState() (MeteredStatus, error) {
	if unLicKey == nil {
		return MeteredStatus{}, errors.New("\u006c\u0069\u0063\u0065ns\u0065\u0020\u006b\u0065\u0079\u0020\u006e\u006f\u0074\u0020\u0073\u0065\u0074")
	}
	if !unLicKey.unknownBoolField1 || len(unLicKey.unknownStrField2) == 0 {
		return MeteredStatus{}, errors.New("\u0061p\u0069 \u006b\u0065\u0079\u0020\u006e\u006f\u0074\u0020\u0073\u0065\u0074")
	}
	_aeg, _bedc := _ege.loadState(unLicKey.unknownStrField2)
	if _bedc != nil {
		logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _bedc)
		return MeteredStatus{}, _bedc
	}
	if _aeg.Docs > 0 {
		_afb := track("", "", true)
		if _afb != nil {
			return MeteredStatus{}, _afb
		}
	}
	mutex.Lock()
	defer mutex.Unlock()
	_dace := _cgf()
	_dace._dac = unLicKey.unknownStrField2
	_fdd, _bedc := _dace.getStatus()
	if _bedc != nil {
		return MeteredStatus{}, _bedc
	}
	if !_fdd.Valid {
		return MeteredStatus{}, errors.New("\u006b\u0065\u0079\u0020\u006e\u006f\u0074\u0020\u0076\u0061\u006c\u0069\u0064")
	}
	_efgd := MeteredStatus{OK: true, Credits: _fdd.OrgCredits, Used: _fdd.OrgUsed}
	return _efgd, nil
}

var _aa *rsa.PublicKey

func (lk *LicenseKey) IsLicensed() bool {
	if lk == nil {
		return false
	}
	return lk.Tier != LicenseTierUnlicensed || lk.unknownBoolField1
}
func _cgf() *meteredClient {
	_gec := meteredClient{baseURL: "h\u0074\u0074\u0070\u0073\u003a\u002f/\u0063\u006c\u006f\u0075\u0064\u002e\u0075\u006e\u0069d\u006f\u0063\u002ei\u006f/\u0061\u0070\u0069", httpClient: &http.Client{Timeout: 30 * time.Second}}
	if _gddg := os.Getenv("\u0055N\u0049\u0044\u004f\u0043_\u004c\u0049\u0043\u0045\u004eS\u0045_\u0053E\u0052\u0056\u0045\u0052\u005f\u0055\u0052L"); strings.HasPrefix(_gddg, "\u0068\u0074\u0074\u0070") {
		_gec.baseURL = _gddg
	}
	return &_gec
}

type MeteredStatus struct {
	OK      bool
	Credits int64
	Used    int64
}

func (_dgf *meteredClient) checkinUsage(form meteredUsageCheckinForm) (meteredUsageCheckinResp, error) {
	form.Package = "\u0075n\u0069\u006f\u0066\u0066\u0069\u0063e"
	form.PackageVersion = common.Version
	var _cff meteredUsageCheckinResp
	_gc := _dgf.baseURL + "\u002f\u006d\u0065\u0074er\u0065\u0064\u002f\u0075\u0073\u0061\u0067\u0065\u005f\u0063\u0068\u0065\u0063\u006bi\u006e"
	_ecg, _dfga := json.Marshal(form)
	if _dfga != nil {
		return _cff, _dfga
	}
	_cga, _dfga := gzipData(_ecg)
	if _dfga != nil {
		return _cff, _dfga
	}
	_bgd, _dfga := http.NewRequest("\u0050\u004f\u0053\u0054", _gc, _cga)
	if _dfga != nil {
		return _cff, _dfga
	}
	_bgd.Header.Add("\u0043\u006f\u006et\u0065\u006e\u0074\u002d\u0054\u0079\u0070\u0065", "\u0061\u0070p\u006c\u0069\u0063a\u0074\u0069\u006f\u006e\u002f\u006a\u0073\u006f\u006e")
	_bgd.Header.Add("\u0043\u006fn\u0074\u0065\u006et\u002d\u0045\u006e\u0063\u006f\u0064\u0069\u006e\u0067", "\u0067\u007a\u0069\u0070")
	_bgd.Header.Add("\u0041c\u0063e\u0070\u0074\u002d\u0045\u006e\u0063\u006f\u0064\u0069\u006e\u0067", "\u0067\u007a\u0069\u0070")
	_bgd.Header.Add("\u0058-\u0041\u0050\u0049\u002d\u004b\u0045Y", _dgf._dac)
	_aca, _dfga := _dgf.httpClient.Do(_bgd)
	if _dfga != nil {
		return _cff, _dfga
	}
	defer _aca.Body.Close()
	if _aca.StatusCode != 200 {
		return _cff, fmt.Errorf("\u0066\u0061i\u006c\u0065\u0064\u0020t\u006f\u0020c\u0068\u0065\u0063\u006b\u0069\u006e\u002c\u0020s\u0074\u0061\u0074\u0075\u0073\u0020\u0063\u006f\u0064\u0065\u0020\u0069s\u003a\u0020\u0025\u0064", _aca.StatusCode)
	}
	_dda, _dfga := _gbc(_aca)
	if _dfga != nil {
		return _cff, _dfga
	}
	_dfga = json.Unmarshal(_dda, &_cff)
	if _dfga != nil {
		return _cff, _dfga
	}
	return _cff, nil
}
func _afce() (net.IP, error) {
	_daf, _eac := net.Dial("\u0075\u0064\u0070", "\u0038\u002e\u0038\u002e\u0038\u002e\u0038\u003a\u0038\u0030")
	if _eac != nil {
		return nil, _eac
	}
	defer _daf.Close()
	_ada := _daf.LocalAddr().(*net.UDPAddr)
	return _ada.IP, nil
}
func _aegb() string {
	_cea := os.Getenv("\u0048\u004f\u004d\u0045")
	if len(_cea) == 0 {
		_cea, _ = os.UserHomeDir()
	}
	return _cea
}
func _efe(_gbe string, _fc string) ([]byte, error) {
	var (
		_fcc int
		_dff string
	)
	for _, _dff = range []string{"\u000a\u002b\u000a", "\u000d\u000a\u002b\r\u000a", "\u0020\u002b\u0020"} {
		if _fcc = strings.Index(_fc, _dff); _fcc != -1 {
			break
		}
	}
	if _fcc == -1 {
		return nil, fmt.Errorf("\u0069\u006e\u0076al\u0069\u0064\u0020\u0069\u006e\u0070\u0075\u0074\u002c \u0073i\u0067n\u0061t\u0075\u0072\u0065\u0020\u0073\u0065\u0070\u0061\u0072\u0061\u0074\u006f\u0072")
	}
	_cg := _fc[:_fcc]
	_bf := _fcc + len(_dff)
	_baa := _fc[_bf:]
	if _cg == "" || _baa == "" {
		return nil, fmt.Errorf("\u0069n\u0076\u0061l\u0069\u0064\u0020\u0069n\u0070\u0075\u0074,\u0020\u006d\u0069\u0073\u0073\u0069\u006e\u0067\u0020or\u0069\u0067\u0069n\u0061\u006c \u006f\u0072\u0020\u0073\u0069\u0067n\u0061\u0074u\u0072\u0065")
	}
	_ccf, _fag := base64.StdEncoding.DecodeString(_cg)
	if _fag != nil {
		return nil, fmt.Errorf("\u0069\u006e\u0076\u0061li\u0064\u0020\u0069\u006e\u0070\u0075\u0074\u0020\u006f\u0072\u0069\u0067\u0069\u006ea\u006c")
	}
	_de, _fag := base64.StdEncoding.DecodeString(_baa)
	if _fag != nil {
		return nil, fmt.Errorf("\u0069\u006e\u0076al\u0069\u0064\u0020\u0069\u006e\u0070\u0075\u0074\u0020\u0073\u0069\u0067\u006e\u0061\u0074\u0075\u0072\u0065")
	}
	_af, _ := pem.Decode([]byte(_gbe))
	if _af == nil {
		return nil, fmt.Errorf("\u0050\u0075\u0062\u004b\u0065\u0079\u0020\u0066\u0061\u0069\u006c\u0065\u0064")
	}
	_cf, _fag := x509.ParsePKIXPublicKey(_af.Bytes)
	if _fag != nil {
		return nil, _fag
	}
	_fagg := _cf.(*rsa.PublicKey)
	if _fagg == nil {
		return nil, fmt.Errorf("\u0050u\u0062\u004b\u0065\u0079\u0020\u0063\u006f\u006e\u0076\u0065\u0072s\u0069\u006f\u006e\u0020\u0066\u0061\u0069\u006c\u0065\u0064")
	}
	_abe := sha512.New()
	_abe.Write(_ccf)
	_eba := _abe.Sum(nil)
	_fag = rsa.VerifyPKCS1v15(_fagg, crypto.SHA512, _eba, _de)
	if _fag != nil {
		return nil, _fag
	}
	return _ccf, nil
}

var _efg = time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC)

func init() {
	_acgg := os.Getenv(UNIOFFICE_LICENSE_PATH)
	_caba := os.Getenv(UNIOFFICE_CUSTOMER_NAME)
	if len(_acgg) == 0 || len(_caba) == 0 {
		return
	}
	_bedd, _egea := ioutil.ReadFile(_acgg)
	if _egea != nil {
		logger.Log.Error("\u0055\u006eab\u006c\u0065\u0020t\u006f\u0020\u0072\u0065ad \u006cic\u0065\u006e\u0073\u0065\u0020\u0063\u006fde\u0020\u0066\u0069\u006c\u0065\u003a\u0020%\u0076", _egea)
		return
	}
	_egea = SetLicenseKey(string(_bedd), _caba)
	if _egea != nil {
		logger.Log.Error("\u0055\u006e\u0061b\u006c\u0065\u0020\u0074o\u0020\u006c\u006f\u0061\u0064\u0020\u006ci\u0063\u0065\u006e\u0073\u0065\u0020\u0063\u006f\u0064\u0065\u003a\u0020\u0025\u0076", _egea)
		return
	}
}
func MakeUnlicensedKey() *LicenseKey {
	lk := LicenseKey{}
	lk.CustomerName = "Unlicensed"
	lk.Tier = LicenseTierUnlicensed
	lk.CreatedAt = time.Now().UTC()
	lk.CreatedAtInt = lk.CreatedAt.Unix()
	return &lk
}

const (
	LicenseTierUnlicensed = "\u0075\u006e\u006c\u0069\u0063\u0065\u006e\u0073\u0065\u0064"
	LicenseTierCommunity  = "\u0063o\u006d\u006d\u0075\u006e\u0069\u0074y"
	LicenseTierIndividual = "\u0069\u006e\u0064\u0069\u0076\u0069\u0064\u0075\u0061\u006c"
	LicenseTierBusiness   = "\u0062\u0075\u0073\u0069\u006e\u0065\u0073\u0073"
)

func (_badg *LicenseKey) ToString() string {
	if _badg.unknownBoolField1 {
		return "M\u0065t\u0065\u0072\u0065\u0064\u0020\u0073\u0075\u0062s\u0063\u0072\u0069\u0070ti\u006f\u006e"
	}
	_eec := fmt.Sprintf("\u004ci\u0063e\u006e\u0073\u0065\u0020\u0049\u0064\u003a\u0020\u0025\u0073\u000a", _badg.LicenseId)
	_eec += fmt.Sprintf("\u0043\u0075s\u0074\u006f\u006de\u0072\u0020\u0049\u0064\u003a\u0020\u0025\u0073\u000a", _badg.CustomerId)
	_eec += fmt.Sprintf("\u0043u\u0073t\u006f\u006d\u0065\u0072\u0020N\u0061\u006de\u003a\u0020\u0025\u0073\u000a", _badg.CustomerName)
	_eec += fmt.Sprintf("\u0054i\u0065\u0072\u003a\u0020\u0025\u0073\n", _badg.Tier)
	_eec += fmt.Sprintf("\u0043r\u0065a\u0074\u0065\u0064\u0020\u0041\u0074\u003a\u0020\u0025\u0073\u000a", common.UtcTimeFormat(_badg.CreatedAt))
	if _badg.ExpiresAt.IsZero() {
		_eec += "\u0045x\u0070i\u0072\u0065\u0073\u0020\u0041t\u003a\u0020N\u0065\u0076\u0065\u0072\u000a"
	} else {
		_eec += fmt.Sprintf("\u0045x\u0070i\u0072\u0065\u0073\u0020\u0041\u0074\u003a\u0020\u0025\u0073\u000a", common.UtcTimeFormat(_badg.ExpiresAt))
	}
	_eec += fmt.Sprintf("\u0043\u0072\u0065\u0061\u0074\u006f\u0072\u003a\u0020\u0025\u0073\u0020<\u0025\u0073\u003e\u000a", _badg.CreatorName, _badg.CreatorEmail)
	return _eec
}
func (_acb defaultStateHolder) loadState(_bdg string) (reportState, error) {
	_cdf := _aegb()
	if len(_cdf) == 0 {
		return reportState{}, errors.New("\u0068\u006fm\u0065\u0020\u0064i\u0072\u0020\u006e\u006f\u0074\u0020\u0073\u0065\u0074")
	}
	_bcg := filepath.Join(_cdf, "\u002eu\u006e\u0069\u0064\u006f\u0063")
	_fde := os.MkdirAll(_bcg, 0777)
	if _fde != nil {
		return reportState{}, _fde
	}
	if len(_bdg) < 20 {
		return reportState{}, errors.New("i\u006e\u0076\u0061\u006c\u0069\u0064\u0020\u006b\u0065\u0079")
	}
	_bec := []byte(_bdg)
	_bga := sha512.Sum512_256(_bec[:20])
	_ggb := hex.EncodeToString(_bga[:])
	_afc := filepath.Join(_bcg, _ggb)
	_fce, _fde := ioutil.ReadFile(_afc)
	if _fde != nil {
		if os.IsNotExist(_fde) {
			return reportState{}, nil
		}
		logger.Log.Error("\u0045R\u0052\u004f\u0052\u003a\u0020\u0025v", _fde)
		return reportState{}, errors.New("\u0069\u006e\u0076a\u006c\u0069\u0064\u0020\u0064\u0061\u0074\u0061")
	}
	const _ccfc = "\u0068\u00619\u004e\u004b\u0038]\u0052\u0062\u004c\u002a\u006d\u0034\u004c\u004b\u0057"
	_fce, _fde = _eae([]byte(_ccfc), _fce)
	if _fde != nil {
		return reportState{}, _fde
	}
	var _cfb reportState
	_fde = json.Unmarshal(_fce, &_cfb)
	if _fde != nil {
		logger.Log.Error("\u0045\u0052\u0052OR\u003a\u0020\u0049\u006e\u0076\u0061\u006c\u0069\u0064\u0020\u0064\u0061\u0074\u0061\u003a\u0020\u0025\u0076", _fde)
		return reportState{}, errors.New("\u0069\u006e\u0076a\u006c\u0069\u0064\u0020\u0064\u0061\u0074\u0061")
	}
	return _cfb, nil
}
func _bad(_gd string, _dd string, _ded string) (string, error) {
	_ebag := strings.Index(_ded, _gd)
	if _ebag == -1 {
		return "", fmt.Errorf("\u0068\u0065a\u0064\u0065\u0072 \u006e\u006f\u0074\u0020\u0066\u006f\u0075\u006e\u0064")
	}
	_gdd := strings.Index(_ded, _dd)
	if _gdd == -1 {
		return "", fmt.Errorf("\u0066\u006fo\u0074\u0065\u0072 \u006e\u006f\u0074\u0020\u0066\u006f\u0075\u006e\u0064")
	}
	_abg := _ebag + len(_gd) + 1
	return _ded[_abg : _gdd-1], nil
}
func (_egc *LicenseKey) TypeToString() string {
	if _egc.unknownBoolField1 {
		return "M\u0065t\u0065\u0072\u0065\u0064\u0020\u0073\u0075\u0062s\u0063\u0072\u0069\u0070ti\u006f\u006e"
	}
	if _egc.Tier == LicenseTierUnlicensed {
		return "\u0055\u006e\u006c\u0069\u0063\u0065\u006e\u0073\u0065\u0064"
	}
	if _egc.Tier == LicenseTierCommunity {
		return "\u0041\u0047PL\u0076\u0033\u0020O\u0070\u0065\u006e\u0020Sou\u0072ce\u0020\u0043\u006f\u006d\u006d\u0075\u006eit\u0079\u0020\u004c\u0069\u0063\u0065\u006es\u0065"
	}
	if _egc.Tier == LicenseTierIndividual || _egc.Tier == "\u0069\u006e\u0064i\u0065" {
		return "\u0043\u006f\u006dm\u0065\u0072\u0063\u0069a\u006c\u0020\u004c\u0069\u0063\u0065\u006es\u0065\u0020\u002d\u0020\u0049\u006e\u0064\u0069\u0076\u0069\u0064\u0075\u0061\u006c"
	}
	return "\u0043\u006fm\u006d\u0065\u0072\u0063\u0069\u0061\u006c\u0020\u004c\u0069\u0063\u0065\u006e\u0073\u0065\u0020\u002d\u0020\u0042\u0075\u0073\u0069ne\u0073\u0073"
}
func (_agfa defaultStateHolder) updateState(_ga, _acg, _bc string, _bbg int, _fgd bool, _aefa int, _bba int, _feaa time.Time, _bda map[string]int) error {
	_bfga := _aegb()
	if len(_bfga) == 0 {
		return errors.New("\u0068\u006fm\u0065\u0020\u0064i\u0072\u0020\u006e\u006f\u0074\u0020\u0073\u0065\u0074")
	}
	_cbef := filepath.Join(_bfga, "\u002eu\u006e\u0069\u0064\u006f\u0063")
	_adb := os.MkdirAll(_cbef, 0777)
	if _adb != nil {
		return _adb
	}
	if len(_ga) < 20 {
		return errors.New("i\u006e\u0076\u0061\u006c\u0069\u0064\u0020\u006b\u0065\u0079")
	}
	_fbge := []byte(_ga)
	_fgc := sha512.Sum512_256(_fbge[:20])
	_fef := hex.EncodeToString(_fgc[:])
	_bfa := filepath.Join(_cbef, _fef)
	var _ceg reportState
	_ceg.Docs = int64(_bbg)
	_ceg.NumErrors = int64(_bba)
	_ceg.LimitDocs = _fgd
	_ceg.RemainingDocs = int64(_aefa)
	_ceg.LastWritten = time.Now().UTC()
	_ceg.LastReported = _feaa
	_ceg.Instance = _acg
	_ceg.Next = _bc
	_ceg.Usage = _bda
	_gdc, _adb := json.Marshal(_ceg)
	if _adb != nil {
		return _adb
	}
	const _abf = "ha9NK8]RbL*m4LKW"
	_gdc, _adb = _cge([]byte(_abf), _gdc)
	if _adb != nil {
		return _adb
	}
	_adb = ioutil.WriteFile(_bfa, _gdc, 0600)
	if _adb != nil {
		return _adb
	}
	return nil
}
func (_da *LicenseKey) getExpiryDateToCompare() time.Time {
	if _da.Trial {
		return time.Now().UTC()
	}
	return common.ReleasedAt
}

var gMap1 map[string]int
