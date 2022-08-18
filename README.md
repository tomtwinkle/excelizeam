# excelizeam
Wrapper to facilitate use of `excelize.StreamWriter` in [qax-os/excelize](https://github.com/qax-os/excelize)

## Motivation

`excelize.StreamWriter` in [qax-os/excelize](https://github.com/qax-os/excelize) has been found to be able to [write large amounts of data at high speed while reducing memory usage](https://xuri.me/excelize/ja/performance.html), and should be used aggressively, but its use is [severely limited](https://pkg.go.dev/github.com/xuri/excelize/v2#File.NewStreamWriter), and [Excel itself will be damaged](https://github.com/qax-os/excelize/issues/1202) if the order of writes is incorrect.

Therefore, a Wrapper was created to automatically meet the restrictions.

[qax-os/excelize](https://github.com/qax-os/excelize) の `excelize.StreamWriter` は大量のデータをメモリ使用量を抑えつつ高速に書き込むことが出来る事が[示されて](https://xuri.me/excelize/ja/performance.html)おり積極的に利用したいが、[利用する際の制限が厳しく](https://pkg.go.dev/github.com/xuri/excelize/v2#File.NewStreamWriter)書き込みの順番を間違えると[Excel自体が破損してしまう](https://github.com/qax-os/excelize/issues/1202)ため、そのままでは使い所が難しい。

そのため制限を自動的に満たすためのWrapperを作成した。

## Usage

TODO
