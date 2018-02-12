package helper

import (
	"io/ioutil"
	"log"
	"os"
	"strings"
)

/**
 * reads license info from a file in the form:
 * <License Name>
 * <License Key>
 */
func ReadLicense(licensePath string) (string, string) {
	if "" == licensePath {
		licensePath = "./license.key"
	}

	if _, err := os.Stat(licensePath); os.IsNotExist(err) {
		return "", ""
	}

	data, err := ioutil.ReadFile(licensePath)

	if nil != err {
		log.Fatal(err)
	}

	parts := strings.Split(string(data), "\n")
	return strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1])
}
