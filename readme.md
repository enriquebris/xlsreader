## example

```go
package main

import (
	"fmt"
	"github.com/enriquebris/xlsreader"
)

func main() {
	reader, err := xlsreader.NewExcelizeReader("myExcelFile.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer reader.Close()

	// read all columns for sheet Furniture
	if err = reader.ReadColumns("Furniture", 6); err != nil {
		fmt.Println(err)
		return
	}

	// print value for productId column, row 10 (Furniture sheet)
	fmt.Println(reader.GetValue("Furniture", "productId", 10))
}
```