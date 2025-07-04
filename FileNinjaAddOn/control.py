from openpyxl.utils import column_index_from_string
import os
from FileNinjaSuite.FileNinjaAddOn.defs import *
from FileNinjaSuite.FileNinjaAddOn.aiScript import queryAI
import traceback


def launchController(excelPath, wb, wsName, aiProcedureName, inColLetter, outColLetter, firstRow, lastRow):
    # set some variables
    baseFilename = os.path.basename(excelPath)
    filenameSansExt, ext = os.path.splitext(baseFilename)
    ws = wb[wsName]
    inColIndex = column_index_from_string(inColLetter)
    outColIndex = column_index_from_string(outColLetter)
    aiProcedureObject = AI_PROCEDURES[aiProcedureName]

    # row error handling
    if lastRow < firstRow:
        return

    # compile list of original input
    inputList = []
    for row in ws.iter_rows(min_row=firstRow, max_row=lastRow, min_col=inColIndex, max_col=inColIndex):
        inputList.append(row[0].value)

    ## prune both sides of inputList to not have None
    # prune from front
    while inputList and inputList[0] is None:
        inputList = inputList[1:]
        firstRow += 1
        break
    # prune from back
    while inputList and inputList[-1] is None:
        inputList = inputList[:-1]
        lastRow -= 1
    if not inputList:
        return

    # query AI (if a prompt was given)
    flaggedChunks = []
    try:
        if aiProcedureObject.promptFile:
            promptFilePath = f"Prompts\\{aiProcedureObject.promptFile}"

            if not os.path.exists(promptFilePath):
                return (-101, promptFilePath)

            with open(promptFilePath, "r", encoding="utf-8") as f:
                developerInstructions = f.read()
            
            #print(len(inputList))    
            
            # Now, chunk into bite-sized queries
            outputList = []
            inputListNumItemsLeft = len(inputList)
            currentIndex = 0
            chunkSize = 100
            #print(f"TOTAL ITEMS: {inputListNumItemsLeft}")
            while inputListNumItemsLeft > 0:
                if inputListNumItemsLeft > chunkSize:
                    aiOutput = queryAI(f"{chunkSize} {str(inputList[currentIndex : currentIndex+chunkSize])}", developerInstructions).output_parsed.correctedFileNamesList
                    aiOutputLength = len(aiOutput)
                    if aiOutputLength == chunkSize:
                        outputList.extend(aiOutput)
                    elif aiOutputLength > chunkSize:
                        #print(f"Outputting {aiOutputLength} items")
                        flaggedChunks.append((currentIndex + firstRow, currentIndex + chunkSize -1 + firstRow))
                        outputList.extend(aiOutput[0:chunkSize])
                    else: # len(aiOutput) < chunkSize
                        #print(f"Outputting {aiOutputLength} items")
                        flaggedChunks.append((currentIndex + firstRow, currentIndex + chunkSize -1 + firstRow))
                        outputList.extend(aiOutput)
                        outputList.extend([""] * (chunkSize - aiOutputLength))

                    currentIndex += chunkSize
                    inputListNumItemsLeft -= chunkSize
                    #print(f"FINISHED: {chunkSize}. ITEMS LEFT: {inputListNumItemsLeft}")
                else:
                    outputList.extend(
                        queryAI(f"{inputListNumItemsLeft} {str(inputList[currentIndex : currentIndex+inputListNumItemsLeft])}", developerInstructions).output_parsed.correctedFileNamesList)
                    #print(f"FINISHED: {inputListNumItemsLeft}. ITEMS LEFT: {0}")
                    currentIndex += inputListNumItemsLeft
                    inputListNumItemsLeft = 0
                    break
            #print(f"PRINTING OUTPUTLIST: {outputList}")
            #print(len(outputList.correctedFileNamesList))
        else:
            outputList = None

        aiProcedureObject.mainFunction(inColIndex, outColIndex, firstRow, lastRow, inputList, outputList, ws)
    except:
        return (STATUS_UNEXPECTED, f"{traceback.format_exc()}")

    # Create RESULTS directory if it does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # save the file and open
    revisedFilename = f"{RESULTS_DIRECTORY}\\{filenameSansExt}-FN-AddOn-Output{ext}"
    wb.save(revisedFilename)
    os.startfile(revisedFilename)
    
    if not flaggedChunks:
        return (STATUS_SUCCESSFUL, None)
    else:
        return (-102, str(flaggedChunks))
