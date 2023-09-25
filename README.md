# exportExcel
导出excel
理论上支持无限维度的标题和数据。
you need Laravel and phpoffice
How to Use?
    ```$file = ExportExcel::explorExcel($exportTitle = [], $exportData = [], $fileInfo = [])```
Please convert your title and data into a specific format
You can see the sample title ```$exportTitle``` and sample data ```$exportData```
本人比较菜所以没考虑服务器性能开销,能用就行(#^.^#)。
I am not very good at it, so I didn't consider the server performance overhead. It works fine for me(#^.^#)