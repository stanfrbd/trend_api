# trend_api

This script helps to get the list of computers and their agent status using Trend Micro CloudOne API in an easy way.

For a better understanding, refer to the official Trend Micro CloudOne API documentation: https://cloudone.trendmicro.com/docs/workload-security/api-reference/

## Install dependencies

```
pip install -r requirements.txt
```

## Usage

```
usage: trend_api.py [-h] {computers}

List computers

positional arguments:
  {computers}  Action to perform

options:
  -h, --help   show this help message and exit
```

# Quick start

Follow the instructions to get an API Key:

## Edit the config file
```
cp secrets-sample.json secrets.json
```

Then edit the config with the good values.

| Secret | Explaination |
|----------|--------------|
|`apiKey`| Your API Key |
|`proxyUrl`| Your proxy if needed - leave blank if not |

```
{
    "apiKey": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
    "proxyUrl": ""
}
```

## Execute the script

> Windows

```
PS C:\Users\Me\Test> python trend_api.py computers
```

> Linux

```
$ python3 trend_api.py computers
```

> Linux  

```
$ chmod +x trend_api.py
$ ./trend_api.py computers
```

## Output

- An Excel (autofiltered) file will be created with datetime.

```
PS C:\Users\Me\Test> python .\trend_api computers
0 agents are active, 49 agents have warnings, 63 agents have errors.
Processed 110 API results.
Successfully created 2023-12-27-14_56_58-trend-api-results.xlsx
```

# Errors

If the `secrets.json` is not properly filled.
```
General error: HTTP error: check your secrets file.
```
