# Fast Gemini AI for Excel

⚡ Excel Macros and Plugins for Google Gemini 🤖 — Now with Ultra-Fast Parallel Processing 🚀 Custom AI functions right in your spreadsheet. Smarter, faster, better.

Get AI-enriched results in milliseconds over hundreds of cells! Automate and systematize various tasks from freeform text cells such as labelling, classification, keywords extraction, sentiment analysis, etc.

This project is a bring your own keys set of Excel Add-in and user-defined functions (UDF) built for integrating with the Gemini AI models by Google. Google offers free API keys but you can use paid API keys for faster results and keeping your data private, see below for more instructions.

If you want the fastest, most up-to-date experience, use the parallel processing `Gemini_udf_p()` user-defined function, see below for more instructions.

## Google Gemini API key generation

To use this software, you need to input your own Gemini API key.

You can generate Gemini API keys here (as of 2025):

https://aistudio.google.com/app/apikey

You can generate free API keys, but they are then rate limited (so you will get some cells with a rate limit message instead of the result you want if you are calculating too many cells at once) and Google will collect all the data you send to train on it.

If you want to lift rate limits to process a lot of cells very fast, as well as keeping your data private, out of Google's training dataset, and have no rate limit, you can generate paid API keys on the same page after activating your Google Cloud billing account and have an active payment method. Note that this also allows you to enjoy free services such as the web interface of Google AI Studio for free but still keep your data [out of the training dataset](https://discuss.ai.google.dev/t/google-ai-studio-is-unsafe-for-private-data/78277/7).

Note that the speed of the response depends on whether the model needs to think to process the request, and also the length of the prompt. Usually, a <1k tokens prompt including other cells contents as context does not take more than a few milliseconds to process for hundreds of cells in parallel.

## Instructions for the Excel Add-in

The Excel Add-in is the most feature-rich, but the source code is unfortunately locked behind a password.

For installation and usage, refer to the instructions by clicking on the link below:  
[Instructions Link](https://www.listendata.com/2023/12/integrate-gemini-into-excel.html)

Note that the Add-in was moved to the `addin` folder. The rest of the instructions remain the same.

## Instructions for the user-defined functions (UDFs)

In the `udf` folder, you can find the user-defined function (UDF) version of the `Gemini()` function in the add-in. This is the only function that was implemented in the UDFs, as it is the most general-purpose, all the other functions behaviors can be reproduced with custom prompts enriched with cells data.

There are two versions of the UDFs:

* `Gemini_udf()` which is a very similar function to the add-in, it is likewise sequential and blocking (ie, when the fill handle is used, a calculation starts and each cell is filled one after the other).

* `Gemini_udf_p()` which is a non-blocking parallel processing ready version, which allows to process with parallelism in a non-blocking way, so that it is much faster to work on multiple cells. It also extremely responsive since the UI is not blocked while the cells are calculating. It is also triggered on-demand via a macro `StartGeminiPoller`, instead of automatically on fill handle activation. It also automatically supports a soft exponential backoff retry mechanism (if hitting rate limit, it will retry with a randomly chosen delay in an increasingly multiplied time window for each retry - both the max delay in milliseconds and the number of retries are configurable as parameters).

Why use user-defined functions instead of the add-in? First they are more lightweight, easier to modify and track changes, and also it seems they are more stable, causing no crashes whereas the add-in may do so.

Note: with a paid API key, getting heavily charged can happen very fast when processing over a lot of cells with the parallel processing function, since it is very fast as it sends requests in parallel, so be careful to test first on a couple of cells before applying to more than hundreds of cells, otherwise you may have a high bill if you have to run multiple times on thousands of cells because your prompt was not as accurate as you thought. So iterate on your prompt first on a few cells before extending to more to keep the cost in your control.

There are two ways to install the UDF functions: the easy method, and the manual method.

### Easy install method

Use this option if you just want to try or if you want to start a new spreadsheet.

Just go to the GitHub Releases and download the `.xslm` file for the latest release. This is a `.xslx` Excel file with all the required modules imported and with cells preconfigured to ease usage.

### Manual install method

Use this option if you want to add the UDFs on your already existing spreadsheet. You will get all the features, but the process is just a bit more involved the first time (but once you know how to do it, it is very fast).

Note: This unfortunately cannot be automated and the online store is only for add-ins, not for UDFs.

#### How to install the UDF functions: preliminary steps

First you need to git clone this repository, do not just download the files, otherwise the line endings will be converted into LF. If you do download the files manually without git clone, then ensure to convert the line endings into CRLF (this can be done in one click with the opensource editor Notepad++).

The functions were tested on Microsoft Office 2016 version 2412 build 18324.20194.

#### How to install Gemini_udf (sequential processing UDF)

To use it, you need to import the .blas file as a module in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11`)

2.  Then go to `File > Import File...` and select `Gemini_udf/geminiAI_Excel_udf.bas`

3. In any cell, type `Gemini_udf("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell can be copied into other cells to calculate in batches and with auto incrementing cell reference.

Note that if you make any change to the sourcecode files, you will need to delete the modules in your Excel sheet and import again the files.

#### How to install Gemini_udf_p (parallel processing UDF)

To use it, you need to import the .bas and .cls files as modules in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11` -- or on some laptops, you may need to press `Fn + Alt + F11`)

2.  Then go to `File > Import File...` and select `Gemini_udf/GeminiAI.bas`. It should be imported under a `Module` folder. NOTE: if you get a #NAME error, this means the module could not be imported, so make sure this file has CRLF line endings (you need to clone the git repository, you cannot just download the raw file from rawgithub).

3.  Do the same with the `Gemini_udf/cGeminiRequest.cls` file, it should be imported under a `Class Module` folder. NOTE: if it gets imported as a simple Module instead of a Class Module (or if you get the error `Expected: instruction end`, then ensure the .cls file has CRLF line returns and not just LF.

4.  Also import `JsonConverter.bas` and `Dictionary.cls`, which are necessary helper modules to parse LLM's JSON responses, they work on both Windows, Mac and Linux.

5.  In any cell, type `Gemini_udf_p("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell should now show "Pending..."

6.  Use the fill handle to copy while increasing the counter over a range of cells. All the cells should show "Pending...". Note: if you want to be able to re-run the command, it is now time to make a copy in another cell or in a notebad, because the result will replace the formula.

7.  Open the Macros viewer (press `Alt + F8` or `Fn + Alt + F8` on some laptops), then select `StartGeminiPoller`, and tap on the `Execute` button. The cells should now all get filled with their final values after a short time (depending on the the selected AI model latency). Note that the results will replace the formula, so the formula is now lost once the poller is executed and the results are received.

8.  Optional: If you want to be able to call the UDF again next time you open the file without having to import again the modules, you can save your worksheet as a `.xslm` file.

Development note: to test changes you might make to the code, after importing both modules, in the Microsoft Visual Basic for Applications window (ALT+F11), click on Debug > Compile (name of) VBA project. This should highlight any issue. This is necessary if after a change you get cells showing a #VALUE error.

## Usage

### `Gemini_udf_p(prompt, api_key, [model], [word_count], [maxDelayMs], [retries], [server_url], [asynchronous])`

This User-Defined Function (UDF) for Microsoft Excel allows users to send prompts directly to Large Language Models (LLMs) from an Excel cell. It supports Google Gemini (bring your own API key) or ChatGPT-compatible API endpoints (including offline self-hosted servers such as ollama). It is designed to operate efficiently, supporting by default asynchronous parallel processing with automatic soft exponential backoff retries (ie, longer and longer wait times after each failures) to avoid hitting rate limits.

**Arguments:**

*   **`prompt`** (String, Required): The text prompt or query to be sent to the LLM.
*   **`api_key`** (String, Optional): Your Google Cloud API or OpenAI-compatible key for authenticating requests. Leave empty if using an offline self-hosted server such as ollama.
*   **`model`** (String, Optional, Default: `"gemini-2.5-flash-preview-05-20"`): Specifies the LLM model to be used.
*   **`word_count`** (Long, Optional, Default: `0`): Desired word count for the generated response.
*   **`maxDelayMs`** (Long, Optional, Default: `500`): Maximum delay in milliseconds between retries for API requests. 
    The wait will be random in this bound. Will be multiplied for each retry.
*   **`retries`** (Integer, Optional, Default: `2`): Number of times to retry an API request if it fails.
*   **`server_url`** (String, Optional, Default: `""`): Custom server URL for the Gemini API endpoint.
*   **`asynchronous`** (Boolean, Optional, Default: `True`): Controls execution mode. `True` for asynchronous parallel processing (returns "Pending...", updates later); `False` for synchronous sequential processing (waits for response after each cell's processing, returns directly).

**Returns:**

*   **Variant**: The generated text response from the Gemini LLM. Returns "Pending..." initially in asynchronous mode, then press ALT+F8 and select StartPoller to launch the requests. In case of an error, it may return an error message or a specific error value.

**Example usages:**
   * offline selfhosted ollama call:
     `=Gemini_udf_p("What is the capital of France?";"a";"qwen2.5:1.5b";0;500;2;"http://localhost:11434")`

## Troubleshooting

### Running the UDF function crashes Excel

There are a LOT of reasons why this may happen, here are common fixes and things to check:

* Check you supplied the right number of arguments and at the right positions and of the right types.
    * If there are optional arguments you are skipping, check you simply provide no value at all (eg, `func(;;10)` instead of `func(0;"";10)`), because otherwise Excel tends to cast "empty" values into NaN or other unexpected and unwanted values instead of the one you supplied explicitly.
* Ensure you imported the latest versions of the modules (CTRL+F11 and delete and import them again).
* Stop, then reinitialize the VBA engine (in the CTRL+F11 window, the pause button then the square button).
* Reevaluate the cell (F9) after reimporting the modules. If only this reevaluated cell works later on, you will know that you need to use the fill handle to reapply the reevaluated cell over the range of cells you want, even if the formula is exactly the same.
* If nothing else works, close all Excel files, then open a new blank one, then importe the modules, then try the function. If it works, you know that it is your other Excel files that have an issue.
    * A way to workaround this issue that may be a bit drastic is to save your Excel file as a .xslx temporarily (instead of .xslm) to delete all macros, then you can reimport the modules manually and then save as a new .xslm , this should fix all issues (if the modules' code works fine of course!).

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

The add-in was originally published by deepanshu88 on [GitHub](https://github.com/deepanshu88/geminiAI-Excel/). The blocking UDF was originally published also by Deepanshu88 in [a blog post](https://www.listendata.com/2023/12/integrate-gemini-into-excel.html). This original work was opensourced at commit 1c4d5b72860890f794b5db7ad66aa545c7949fa9 of the original repository, before being deleted. Unfortunately, the addin remains password-protected although technically it is opensource licensed, but the sourcecode is not currently accessible.

The parallel processing UDF was made by Stephen Karl Larroque and first published on [GitHub](https://github.com/lrq3000/geminiAI-Excel).

This project is using the incredible [Tim Hall's VBA-JSON](https://github.com/VBA-tools/VBA-JSON/) and [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary) to parse LLM API's responses reliably.
