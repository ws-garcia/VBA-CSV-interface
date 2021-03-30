---
title: GuessDelimiters
parent: Methods
grand_parent: API
nav_order: 13
---

# GuessDelimiters
{: .d-inline-block }

New
{: .label .label-purple }

Runs an analysis trying to guess delimiters used on the CSV/TSV file indicated in the `.path` property of the configuration object.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`GuessDelimiters`*(confObj)*

### Parameters

The required *confObj* argument is an identifier specifying a `parserConfig` object variable.

### Returns value

_None_

---

## Behavior

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>Only a few CSV/TSV records will be used for guess delimiters. The results of the analysis are saved in the *confObj* parameter, this means the *confObj* object properties will be altered.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)