## User-defined function version

This is the user-defined function UDF) version of the Gemini() function in the add-in.

To use it, you need to import the .blas file as a module in any Excel file where you want to use the function.

1.  Open the **VBA editor** (`Alt + F11`)
    
2.  In Project Explorer, right-click the old module (e.g., `Module1`) → choose **Remove** → **Do not export**
    
3.  Then go to `File > Import File...` and select the updated `.bas`

4. In any cell, type `Gemini_udf("Prompt text"&A1;"Gemini_API_key";"Gemini_model")`. The Gemini model is optional, and also a max word count for the response can be added as a 4th optional argument. The cell can be copied into other cells to calculate in batches and with auto incrementing cell reference.
