---
title: SortingAlgorithms
parent: Enumerations
grand_parent: API
nav_order: 4
---

# SortingAlgorithms Enum
{: .d-inline-block }

New
{: .label .label-purple }

Provides a list of constants to configure the sorting algorithm used when sorting data imported from a CSV file.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|0|*SA_Quicksort*|
|1|*SA_TimSort*|
|2|*SA_HeapSort*|
|3|*SA_MergeSort*|

---

## Syntax

*variable* = `SortingAlgorithms`.*Constant*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The default value for the `SortingAlgorithms` enumeration is `SA_Quicksort` which is a variant of the classic Quicksort.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/sort.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)