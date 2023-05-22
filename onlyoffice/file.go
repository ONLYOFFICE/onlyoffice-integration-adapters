/**
 *
 * (c) Copyright Ascensio System SIA 2023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

// Package onlyoffice provides onlyoffice document server specific utility functions
//
// The onlyoffice package's structures are self-initialized by fx and bootstrapper.
package onlyoffice

import (
	"context"
	"errors"
	"net/http"
	"path/filepath"
	"strconv"
	"strings"
)

var (
	ErrOnlyofficeExtensionNotSupported = errors.New("file extension is not supported")
	ErrInvalidContentLength            = errors.New("could not perform api actions due to exceeding content-length")
)

const (
	_OnlyofficeWordType  string = "word"
	_OnlyofficeCellType  string = "cell"
	_OnlyofficeSlideType string = "slide"
)

// OnlyofficeEditableExtensions maps editable according to the API documentation
// extensions to document types.
var OnlyofficeEditableExtensions map[string]string = map[string]string{
	"docm":  _OnlyofficeWordType,
	"docx":  _OnlyofficeWordType,
	"docxf": _OnlyofficeWordType,
	"oform": _OnlyofficeWordType,
	"dotm":  _OnlyofficeWordType,
	"dotx":  _OnlyofficeWordType,
	"xlsm":  _OnlyofficeCellType,
	"xlsx":  _OnlyofficeCellType,
	"xltm":  _OnlyofficeCellType,
	"xltx":  _OnlyofficeCellType,
	"potm":  _OnlyofficeSlideType,
	"potx":  _OnlyofficeSlideType,
	"ppsm":  _OnlyofficeSlideType,
	"ppsx":  _OnlyofficeSlideType,
	"pptm":  _OnlyofficeSlideType,
	"pptx":  _OnlyofficeSlideType,
}

// OnlyofficeOOXMLEditableExtensions maps convertable OOXML according to the API documentation
// extensions to document types.
var OnlyofficeOOXMLEditableExtensions map[string]string = map[string]string{
	"doc":   _OnlyofficeWordType,
	"dot":   _OnlyofficeWordType,
	"fodt":  _OnlyofficeWordType,
	"mht":   _OnlyofficeWordType,
	"xml":   _OnlyofficeWordType,
	"sxw":   _OnlyofficeWordType,
	"stw":   _OnlyofficeWordType,
	"htm":   _OnlyofficeWordType,
	"mhtml": _OnlyofficeWordType,
	"wps":   _OnlyofficeWordType,
	"wpt":   _OnlyofficeWordType,
	"fods":  _OnlyofficeCellType,
	"xls":   _OnlyofficeCellType,
	"xlt":   _OnlyofficeCellType,
	"sxc":   _OnlyofficeCellType,
	"et":    _OnlyofficeCellType,
	"ett":   _OnlyofficeCellType,
	"xlsb":  _OnlyofficeCellType,
	"fodp":  _OnlyofficeSlideType,
	"pot":   _OnlyofficeSlideType,
	"pps":   _OnlyofficeSlideType,
	"ppt":   _OnlyofficeSlideType,
	"sxi":   _OnlyofficeSlideType,
	"dps":   _OnlyofficeSlideType,
	"dpt":   _OnlyofficeSlideType,
}

// OnlyofficeDataLossEditableExtensions maps not fully editable according to the API documentation
// extensions to document types.
var OnlyofficeDataLossEditableExtensions map[string]string = map[string]string{
	"epub": _OnlyofficeWordType,
	"fb2":  _OnlyofficeWordType,
	"html": _OnlyofficeWordType,
	"odt":  _OnlyofficeWordType,
	"ott":  _OnlyofficeWordType,
	"rtf":  _OnlyofficeWordType,
	"txt":  _OnlyofficeWordType,
	"csv":  _OnlyofficeCellType,
	"ods":  _OnlyofficeCellType,
	"ots":  _OnlyofficeCellType,
	"odp":  _OnlyofficeSlideType,
	"otp":  _OnlyofficeSlideType,
}

// OnlyofficeViewOnlyExtensions maps only read-only according to the API documentation
// file extensions to document types.
var OnlyofficeViewOnlyExtensions map[string]string = map[string]string{
	"djvu": _OnlyofficeWordType,
	"oxps": _OnlyofficeWordType,
	"pdf":  _OnlyofficeWordType,
	"xps":  _OnlyofficeWordType,
}

// An OnlyofficeFileUtility provides basic filename interaction functions.
// The implementation structure is expected to be intialized automatically by fx and bootstrapper.
type OnlyofficeFileUtility interface {
	// ValidateFileSize takes a context, size limit and url to send head request to.
	// It returns an error if file's content-type exceeds the limit passed to the function.
	//
	// A successful ValidateFileSize returns err == nil.
	ValidateFileSize(ctx context.Context, limit int64, url string) error
	// EscapeFilename take a file name and sanitizes it.
	// It returns a sanitized file name.
	//
	// A successful EscapeFilename return a non-empty string.
	EscapeFilename(filename string) string
	// IsExtensionSupported takes file extensions and checks all onlyoffice file
	// maps.
	// It returns true/false flag.
	IsExtensionSupported(fileExt string) bool
	// IsExtensionEditable takes file extension and checks onlyoffice editable map.
	// It returns true/false.
	IsExtensionEditable(fileExt string) bool
	// IsExtensionViewOnly takes file extension and checks onlyoffice view only map.
	// It returns true/false.
	IsExtensionViewOnly(fileExt string) bool
	// IsExtensionLossEditable takes file extension and checks onlyoffice dataloss map.
	// It returns true/false.
	IsExtensionLossEditable(fileExt string) bool
	// IsExtensionOOXMLConvertable take file extension and checks onlyoffice ooxml convertable map.
	// It returns true/false.
	IsExtensionOOXMLConvertable(fileExt string) bool
	// GetFilenameWithoutExtension takes full filename and strips out file extension.
	// It returns file name without extension.
	GetFilenameWithoutExtension(filename string) string
	// GetFileType takes file extension and maps it to onlyoffice file type.
	// It returns file type and the first encountered error.
	//
	// A successful GetFileType returns a non-empty file type and err == nil
	GetFileType(fileExt string) (string, error)
	// GetFileExt take file name and strips out the base of the name, leaving only
	// file extension.
	GetFileExt(filename string) string
}

// An OnlyofficeFileUtility constructor. Called automatically by fx and
// bootstrapper.
//
// Returns an onlyoffice file utility implementation based on configuration.
func NewOnlyofficeFileUtility() OnlyofficeFileUtility {
	return fileUtility{}
}

type fileUtility struct{}

func (u fileUtility) ValidateFileSize(ctx context.Context, limit int64, url string) error {
	resp, err := http.Head(url)

	if err != nil {
		return err
	}

	if val, err := strconv.ParseInt(resp.Header.Get("Content-Length"), 10, 0); val > limit || err != nil {
		return ErrInvalidContentLength
	}

	return nil
}

func (u fileUtility) EscapeFilename(filename string) string {
	f := strings.ReplaceAll(filename, "\\", ":")
	f = strings.ReplaceAll(f, "/", ":")
	return f
}

func (u fileUtility) IsExtensionSupported(fileExt string) bool {
	ext := strings.ToLower(fileExt)
	if _, exists := OnlyofficeDataLossEditableExtensions[ext]; exists {
		return true
	}

	if _, exists := OnlyofficeEditableExtensions[ext]; exists {
		return true
	}

	if _, exists := OnlyofficeOOXMLEditableExtensions[ext]; exists {
		return true
	}

	if _, exists := OnlyofficeViewOnlyExtensions[ext]; exists {
		return true
	}

	return false
}

func (u fileUtility) IsExtensionEditable(fileExt string) bool {
	_, exists := OnlyofficeEditableExtensions[strings.ToLower(fileExt)]
	return exists
}

func (u fileUtility) IsExtensionViewOnly(fileExt string) bool {
	_, exists := OnlyofficeViewOnlyExtensions[strings.ToLower(fileExt)]
	return exists
}

func (u fileUtility) IsExtensionLossEditable(fileExt string) bool {
	_, exists := OnlyofficeDataLossEditableExtensions[strings.ToLower(fileExt)]
	return exists
}

func (u fileUtility) IsExtensionOOXMLConvertable(fileExt string) bool {
	_, exists := OnlyofficeOOXMLEditableExtensions[strings.ToLower(fileExt)]
	return exists
}

func (u fileUtility) GetFilenameWithoutExtension(filename string) string {
	return strings.TrimSuffix(filename, filepath.Ext(filename))
}

func (u fileUtility) GetFileType(fileExt string) (string, error) {
	ext := strings.ToLower(fileExt)
	if fType, exists := OnlyofficeEditableExtensions[ext]; exists {
		return fType, nil
	}

	if fType, exists := OnlyofficeDataLossEditableExtensions[ext]; exists {
		return fType, nil
	}

	if fType, exists := OnlyofficeOOXMLEditableExtensions[ext]; exists {
		return fType, nil
	}

	if fType, exists := OnlyofficeViewOnlyExtensions[ext]; exists {
		return fType, nil
	}

	return "", ErrOnlyofficeExtensionNotSupported
}

func (u fileUtility) GetFileExt(filename string) string {
	return strings.ReplaceAll(filepath.Ext(filename), ".", "")
}
