## User-defined function version

This is the user-defined function (UDF) version of the `Gemini()` function in the add-in.

There are two versions:

* `Gemini_udf()` which is a very similar function to the add-in, it is likewise sequential and blocking (ie, when the fill handle is used, a calculation starts and each cell is filled one after the other).
* `Gemini_udf_p()` which is a non-blocking parallel processing ready version, which allows to process with parallelism in a non-blocking way, so that it is much faster to work on multiple cells. It is also triggered on-demand via a macro `StartGeminiPoller`, instead of automatically on fill handle activation.

### How to use Gemini_udf (sequential processing UDF)

To use it, you need to import the .blas file as a module in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11`)

2.  Then go to `File > Import File...` and select `Gemini_udf/geminiAI_Excel_udf.bas`

3. In any cell, type `Gemini_udf("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell can be copied into other cells to calculate in batches and with auto incrementing cell reference.

Note that if you make any change to the sourcecode files, you will need to delete the modules in your Excel sheet and import again the files.

### How to use Gemini_udf_p (parallel processing UDF)

To use it, you need to import the .bas and .cls files as modules in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11` -- or on some laptops, you may need to press `Fn + Alt + F11`)
    
2.  Then go to `File > Import File...` and select `Gemini_udf/GeminiAI.bas`. It should be imported under a `Module` folder.

3. Do the same with the `Gemini_udf/cGeminiRequest.cls` file, it should be imported under a `Class Module` folder. NOTE: if it gets imported as a simple Module instead of a Class Module (or if you get the error `Expected: instruction end`, then ensure the .cls file has CRLF line returns and not just LF.

4. In any cell, type `Gemini_udf_p("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell should now show "Pending..."

5. Use the fill handle to copy while increasing the counter over a range of cells. All the cells should show "Pending...". Note: if you want to be able to re-run the command, it is now time to make a copy in another cell or in a notebad, because the result will replace the formula.

6. Open the Macros viewer (press `Alt + F8` or `Fn + Alt + F8` on some laptops), then select `StartGeminiPoller`, and tap on the `Execute` button. The cells should now all get filled with their final values after a short time (depending on the the selected AI model latency). Note that the results will replace the formula, so the formula is now lost once the poller is executed and the results are received.

7. Optional: If you want to be able to call the UDF again next time you open the file without having to import again the modules, you can save your worksheet as a `.xslm` file.
