## Excel Plugin for Google's Gemini AI Model

⚡ Excel Plugin for Google Gemini 🤖 — Now with Ultra-Fast Parallel Processing 🚀 Custom AI functions right in your spreadsheet. Smarter, faster, better.

This project is a set of Excel Add-in and user-defined functions (UDF) built for integrating with the Gemini AI models by Google.

## Instructions for the Excel Add-in

The Excel Add-in is the most feature-rich, but the source code is unfortunately locked behind a password.

For installation and usage, refer to the instructions by clicking on the link below:  
[Instructions Link](https://www.listendata.com/2023/12/integrate-gemini-into-excel.html)

Note that the Add-in was moved to the `addin` folder. The rest of the instructions remain the same.

## Instructions for the user-defined functions (UDFs)

In the `udf` folder, you can find the user-defined function (UDF) version of the `Gemini()` function in the add-in. This is the only function that was implemented in the UDFs, as it is the most general-purpose, all the other functions behaviors can be reproduced with custom prompts enriched with cells data.

There are two versions of the UDFs:

* `Gemini_udf()` which is a very similar function to the add-in, it is likewise sequential and blocking (ie, when the fill handle is used, a calculation starts and each cell is filled one after the other).

* `Gemini_udf_p()` which is a non-blocking parallel processing ready version, which allows to process with parallelism in a non-blocking way, so that it is much faster to work on multiple cells. It is also triggered on-demand via a macro `StartGeminiPoller`, instead of automatically on fill handle activation.

### How to install the UDF functions

First you need to git clone this repository, do not just download the files, otherwise the line endings will be converted into LF. If you do download the files manually without git clone, then ensure to convert the line endings into CRLF (this can be done in one click with the opensource editor Notepad++).

The functions were tested on Microsoft Office 2016 version 2412 build 18324.20194.

### How to use Gemini_udf (sequential processing UDF)

To use it, you need to import the .blas file as a module in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11`)

2.  Then go to `File > Import File...` and select `Gemini_udf/geminiAI_Excel_udf.bas`

3. In any cell, type `Gemini_udf("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell can be copied into other cells to calculate in batches and with auto incrementing cell reference.

Note that if you make any change to the sourcecode files, you will need to delete the modules in your Excel sheet and import again the files.

### How to use Gemini_udf_p (parallel processing UDF)

To use it, you need to import the .bas and .cls files as modules in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11` -- or on some laptops, you may need to press `Fn + Alt + F11`)
    
2.  Then go to `File > Import File...` and select `Gemini_udf/GeminiAI.bas`. It should be imported under a `Module` folder. NOTE: if you get a #NAME error, make sure this file has CRLF line endings (you need to clone the git repository, you cannot just download the raw file from rawgithub).

3. Do the same with the `Gemini_udf/cGeminiRequest.cls` file, it should be imported under a `Class Module` folder. NOTE: if it gets imported as a simple Module instead of a Class Module (or if you get the error `Expected: instruction end`, then ensure the .cls file has CRLF line returns and not just LF.

4. In any cell, type `Gemini_udf_p("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell should now show "Pending..."

5. Use the fill handle to copy while increasing the counter over a range of cells. All the cells should show "Pending...". Note: if you want to be able to re-run the command, it is now time to make a copy in another cell or in a notebad, because the result will replace the formula.

6. Open the Macros viewer (press `Alt + F8` or `Fn + Alt + F8` on some laptops), then select `StartGeminiPoller`, and tap on the `Execute` button. The cells should now all get filled with their final values after a short time (depending on the the selected AI model latency). Note that the results will replace the formula, so the formula is now lost once the poller is executed and the results are received.

7. Optional: If you want to be able to call the UDF again next time you open the file without having to import again the modules, you can save your worksheet as a `.xslm` file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

The add-in was originally published by deepanshu88 on [GitHub](https://github.com/deepanshu88/geminiAI-Excel/). The blocking UDF was originally published also by Deepanshu88 in [a blog post](https://www.listendata.com/2023/12/integrate-gemini-into-excel.html). This original work was opensourced at commit 1c4d5b72860890f794b5db7ad66aa545c7949fa9 of the original repository, before being deleted. Unfortunately, the addin remains password-protected although technically it is opensource licensed, but the sourcecode is not currently available.

The parallel processing UDF was made by Stephen Karl Larroque and first published on [GitHub](https://github.com/lrq3000/geminiAI-Excel).
