# excelizeam
Wrapper to facilitate use of `excelize.StreamWriter` in [qax-os/excelize](https://github.com/qax-os/excelize)

## Motivation

`excelize.StreamWriter` in [qax-os/excelize](https://github.com/qax-os/excelize) has been found to be able to [write large amounts of data at high speed while reducing memory usage](https://xuri.me/excelize/en/performance.html), and should be used aggressively, but its use is [severely limited](https://pkg.go.dev/github.com/xuri/excelize/v2#File.NewStreamWriter), and [Excel itself will be damaged](https://github.com/qax-os/excelize/issues/1202) if the order of writes is incorrect.

Therefore, a Wrapper was created to automatically meet the restrictions.

When using SyncMethod such as `SetCellValue()`, there is no advantage in speed, but by using Async function such as `SetCellValueAsync()`, it is possible to achieve almost the same speed as `excelize.StreamWriter`.

[qax-os/excelize](https://github.com/qax-os/excelize) の `excelize.StreamWriter` は大量のデータをメモリ使用量を抑えつつ高速に書き込むことが出来る事が[示されて](https://xuri.me/excelize/ja/performance.html)おり積極的に利用したいが、[利用する際の制限が厳しく](https://pkg.go.dev/github.com/xuri/excelize/v2#File.NewStreamWriter)書き込みの順番を間違えると[Excel自体が破損してしまう](https://github.com/qax-os/excelize/issues/1202)ため、そのままでは使い所が難しい。

そのため制限を自動的に満たすためのWrapperを作成した。

`SetCellValue()` のようなSyncMethodを使用する場合は速度面でのメリットを享受出来ないが、 `SetCellValueAsync()` のようなAsync関数を利用することで `excelize.StreamWriter` とほぼ同様の速度を出すことが可能になる。

```
BenchmarkExcelizeam
BenchmarkExcelizeam/Excelize
BenchmarkExcelizeam/Excelize-12                  	       9	 124964492 ns/op
BenchmarkExcelizeam/Excelize_Async
BenchmarkExcelizeam/Excelize_Async-12   	               6	 208647106 ns/op
BenchmarkExcelizeam/Excelize_StreamWriter
BenchmarkExcelizeam/Excelize_StreamWriter-12         	      12	  89386412 ns/op
BenchmarkExcelizeam/Excelizeam_Sync
BenchmarkExcelizeam/Excelizeam_Sync-12               	       5	 232665514 ns/op
BenchmarkExcelizeam/Excelizeam_Async
BenchmarkExcelizeam/Excelizeam_Async-12              	      13	  86562598 ns/op
```

## Usage

### install your project

```
go get github.com/tomtwinkle/excelizeam@v0.3.4
```

TODO
