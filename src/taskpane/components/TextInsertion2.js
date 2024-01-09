import React, { useState, useEffect } from "react";
import axios from 'axios';
import "./TextInsertion.css"

function TextInsertion() {
  const [updateComment, setUpdatedComment] = useState([]);
  const [allComments, setAllComments] = useState([]);
  const [excelData, setExcelData] = useState([])
  const [activeCellValue, setActiveCellValue] = useState("");
  const [updateC, setUpdateC] = useState(false);
  const [fillAuditData, setFillAuditData] = useState([]);
  const [filterCommentData, setFillterCommentData] = useState([]);
  const [liveCell, setLiveCell] = useState();
  const [isAddComment, setIsAddComment] = useState(false);
  const [showInputFields, setShowInputFields] = useState([]);
  const [customComments, setCustomComments] = useState('');
  const [filterCellComments, setFilterCellComments] = useState([])
  let cellChangeAuthor = "";

  useEffect(() => {

    initializeEventHandlers();
  }, [filterCommentData]);




  async function initializeEventHandlers() {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem("Sheet1");
        const sheetsList = context.workbook.worksheets;
        sheetsList.load("items/name");
        await context.sync();
        //console.log(sheetsList.items)
        sheetsList.items.forEach(function (sheet) {
          console.log(sheet.name);
        });
        //  worksheet.onChanged.add(handleChange);
        let range = worksheet.getRange();
        // range.clear();
        worksheet.onChanged.add(handleChange);

        worksheet.onSelectionChanged.add(getActiveCellData);
        worksheet.comments.onAdded.add(commentAdded);
        // worksheet.comments.onChanged.add(commentChanged);


        await context.sync();
        console.log("Event handler successfully registered for onChanged event in the worksheet.");
      });
    } catch (error) {
      console.error("Error initializing event handlers: " + error.message);
    }
  }



  async function commentChanged(event) {
    console.log(event)
    if (event.changeType === 'ReplyAdded') {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Sheet1");
        const comment = sheet.comments.getItemAt();
        const reply = comment.replies.getItemAt();
        console.log(reply)
        reply.load("content");
        // Sync to load the content of the comment reply.
        await context.sync();

        // Append "Please!" to the end of the comment reply.
        reply.content += " Please!";
        console.log(reply)
        await context.sync();
      });
    }

    else if (event.changeType === 'CommentEdited') {
      try {
        await Excel.run(async (context) => {
          const changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);
          changedComment.load(["content", "authorName", "creationDate"]);

          await context.sync();

          console.log(changedComment)


          // Log the details of the changed comment
          console.log("A comment was changed:");
          console.log(`ID: ${event.commentDetails[0].commentId}`);
          console.log(`Updated comment content: ${changedComment.content}`);
          console.log(`Comment author: ${changedComment.authorName}`);
          console.log(`Creation date: ${changedComment.creationDate}`);
        });
      } catch (error) {
        console.error("Error in commentChanged function:", error.message);
      }
    }

  }

  async function commentAdded(event) {
  
    await Excel.run(async (context) => {
      const addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);
      addedComment.load(["content", "authorName", "creationDate"]);
      await context.sync();
      const activeCell = context.workbook.getActiveCell();
      activeCell.load("address");
      await context.sync();

      const cellAddr = activeCell.address.split("!")[1];

      await context.sync();

      let sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      let commentsDatasheet;
      let commentsDataSheetExists = sheets.items.some(sheet => sheet.name === "CommentsData");

      if (!commentsDataSheetExists) {
        commentsDatasheet = sheets.add("CommentsData");
        let commentsDataRange = commentsDatasheet.getRange("A1:E1");
        let CommentsDataHeaders = [
          ["CellAddress", "ChangeType", "Content", "AuthorName", "CreationDate"]
        ];
        commentsDataRange.values = CommentsDataHeaders;
        commentsDataRange.format.columnWidth = 130;
        await context.sync();
      }

      else {
        commentsDatasheet = sheets.getItem("CommentsData");
      }

      if (addedComment.content !== 'Test1234567') {
        let lastRow;
        let usedRange = commentsDatasheet.getUsedRange().load("rowCount");
        await context.sync();

        if (usedRange.rowCount) {
          lastRow = usedRange.rowCount;
        } else {
          console.log("No existing data in 'CellsData' worksheet.");
        }
        let commentsdataToStore = [
          [cellAddr, event.type, addedComment.content, addedComment.authorName, addedComment.creationDate]
        ];
        let dataRange = commentsDatasheet.getRange("A" + (lastRow + 1) + ":E" + (lastRow + 1));
        dataRange.values = commentsdataToStore;

      }
      
      const newComment = {
        comment: addedComment.content,
        createdBy: addedComment.authorName,
        created: addedComment.creationDate,
      };
      setUpdatedComment([newComment]);
      

    });
  }

  async function getActiveCellData() {
    setUpdatedComment('');
    await Excel.run(async (context) => {
      try {
        const activeCell = context.workbook.getActiveCell();
        activeCell.load("address");
        activeCell.load("values");
        await context.sync();
        const cellAddr = activeCell.address.split("!")[1];
        setLiveCell(cellAddr);
        const sheet2 = context.workbook.worksheets.getItemOrNullObject("CellsData");
        const commentsSheet = context.workbook.worksheets.getItemOrNullObject("CommentsData");
        const cellCommentssSheet = context.workbook.worksheets.getItemOrNullObject("TrxComments");
        await context.sync();
        if (!sheet2.isNullObject) {
          const range = sheet2.getUsedRange();
          range.load("values");
          await context.sync();
          const jsonData = convertSheetDataToJson(range.values);
          const filterAuditHistory = jsonData.filter(item => item.CellAddress === cellAddr);
          setFillAuditData(filterAuditHistory);
        } else {
          console.log("Sheet 'CellsData' not found.");
        }
        if (!commentsSheet.isNullObject) {
          const cRange = commentsSheet.getUsedRange();
          cRange.load("values");
          await context.sync();
          const cJsonData = convertCommentsDataToJson(cRange.values);
          const filterCommentHistory = cJsonData.filter(item => item.CellAddress === cellAddr);
          setFillterCommentData(filterCommentHistory);
        }
        if (!cellCommentssSheet.isNullObject) {
          const cellRange = cellCommentssSheet.getUsedRange();
          cellRange.load("values");
          await context.sync();
          const cellJsonData = convertCellCommentsDataToJson(cellRange.values);
          const filterCellCommentHistory = cellJsonData.filter(item => item.CellAddress === cellAddr);
          console.log(filterCellCommentHistory);
          setFilterCellComments(filterCellCommentHistory);
        }
        setActiveCellValue(activeCell.values[0][0]);
      } catch (error) {
        console.error("Error in getActiveCellData:", error.message);
      }
    });
  }



  async function handleChange(event) {

    await Excel.run(async (context) => {
      let sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      let cellDatasheet;
      let cellsDataSheetExists = sheets.items.some(sheet => sheet.name === "CellsData");


      if (cellsDataSheetExists) {
        cellDatasheet = sheets.getItem("CellsData");
      } else {
        cellDatasheet = sheets.add("CellsData");
        // cellDatasheet.visibility = 'Hidden';
      }

      await addCommentToCell("Sheet1!A2", "Test1234567");

      let cellDataRange = cellDatasheet.getRange("A1:G1");
      let CellDataHeaders = [
        ["ID", "CellAddress", "ChangeType", "ValueBefore", "ValueAfter", "ChangedBy", "TimeStamp"]
      ];
      cellDataRange.values = CellDataHeaders;
      cellDataRange.format.columnWidth = 130;

      let lastRow;

      let usedRange = cellDatasheet.getUsedRange().load("rowCount");
      await context.sync();

      if (usedRange.rowCount) {
        lastRow = usedRange.rowCount;

      } else {
        console.log("No existing data in 'CellsData' worksheet.");

      }


      let beforeVal = event.details.valueBefore !== '' ? event.details.valueBefore : 'null';
      let dataToStore = [
        [lastRow + 1, event.address, event.changeType, beforeVal, event.details.valueAfter, cellChangeAuthor, new Date().toLocaleString()]
      ];

      let dataRange = cellDatasheet.getRange("A" + (lastRow + 1) + ":G" + (lastRow + 1));
      dataRange.values = dataToStore;

      await deleteCommented("Sheet1!A2");
      await getActiveCellData();
      await context.sync();

    }).catch(errorHandlerFunction);
  }

  function convertSheetDataToJson(sheetData) {

    const jsonData = sheetData.map(row => {
      const rowData = {};
      row.forEach((value, index) => {
        // Assuming the first row contains headers
        const header = sheetData[0][index];
        rowData[header] = value;
      });
      return rowData;
    });
    jsonData.shift();

    return jsonData;
  }

  function convertCommentsDataToJson(CsheetData) {
    const commentsjsonData = CsheetData.map(row => {
      const CrowData = {};
      row.forEach((value, index) => {
        const Cheader = CsheetData[0][index];
        CrowData[Cheader] = value;
      });
      return CrowData;
    });
    commentsjsonData.shift();

    return commentsjsonData;
  }

  function convertCellCommentsDataToJson(CellsheetData) {
    const cellCommentsjsonData = CellsheetData.map(row => {
      const CellrowData = {};
      row.forEach((value, index) => {
        const Cellheader = CellsheetData[0][index];
        CellrowData[Cellheader] = value;
      });
      return CellrowData;
    });
    cellCommentsjsonData.shift();

    return cellCommentsjsonData;
  }

  async function addCommentToCell(cellAddress, content) {

    await Excel.run(async (context) => {
      // Add a comment to A2 on the "Sheet1" worksheet.
      let comments = context.workbook.comments;
      let comment = comments.add(cellAddress, content);
      comment.load("authorName");
      await context.sync();
      cellChangeAuthor = comment.authorName
      console.log("Comment added by: " + comment.authorName);
    });
  }

  async function deleteCommented(CellAddr) {
    await Excel.run(async (context) => {
      // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
      context.workbook.comments.getItemByCell(CellAddr).delete();
      await context.sync();
    });
  }
  const showCommentBox = (index) => {
    // console.log(index);
    // Toggle the visibility for the corresponding item
    setShowInputFields((prev) => {
      const updatedVisibility = [...prev];
      updatedVisibility[index] = !prev[index];
      return updatedVisibility;
    });
  };
  function handleCustomCommentChange(e, index) {

    setCustomComments(e.target.value);
  }

  async function submitComment(ID, cellPostion, valueAfter) {
    await Excel.run(async (context) => {
      let sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      let cellDatasheet;
      let cellsDataSheetExists = sheets.items.some(sheet => sheet.name === "TrxComments");
      if (cellsDataSheetExists) {
        cellDatasheet = sheets.getItem("TrxComments");
      } else {
        cellDatasheet = sheets.add("TrxComments");
        let cellDataRange = cellDatasheet.getRange("A1:E1");
        let CellDataHeaders = [
          ["CellAddress", "CellValue", "Comment", "TimeStamp", "CellsDataSheetId"]
        ];
        cellDataRange.values = CellDataHeaders;
        cellDataRange.format.columnWidth = 130;
        // cellDatasheet.visibility = 'Hidden';
      }

      let lastRow;

      let usedRange = cellDatasheet.getUsedRange().load("rowCount");
      await context.sync();

      if (usedRange.rowCount) {
        lastRow = usedRange.rowCount;

      } else {
        console.log("No existing data in 'CellsData' worksheet.");

      }

      let dataToStore = [
        [cellPostion, valueAfter, customComments, new Date().toLocaleString(), ID]
      ];

      let dataRange = cellDatasheet.getRange("A" + (lastRow + 1) + ":E" + (lastRow + 1));
      dataRange.values = dataToStore;


      // await getActiveCellData();
      // setShowInputFields([])
      setCustomComments('')
    }).catch(errorHandlerFunction);
  }

  return (
    <div>
   
      <h1>Audit History</h1>

      {fillAuditData && fillAuditData.length > 0 ? (
        fillAuditData.map((item, index) => (
          <div key={index}>
            <p>{item.ChangedBy} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;{item.TimeStamp}</p>
            <p>
              Value: {item.ValueAfter} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{" "}
              <button onClick={() => showCommentBox(index)}>
                {showInputFields[index] ? "Hide Comment" : "Add Comment"}
              </button> &nbsp; &nbsp; &nbsp; &nbsp;
              {showInputFields[index] && <><input type="text" name="" value={customComments} onChange={(e) => handleCustomCommentChange(e, index)} /> <button onClick={() => submitComment(item.ID, item.CellAddress, item.ValueAfter)}>submit</button></>}

            </p>
            {
              filterCellComments?.map((cellcomment) => (
                cellcomment.CellsDataSheetId === item.ID ? (
                  <div key={cellcomment.ID}>
                    <p>{cellcomment.Comment}</p>
                  </div>
                ) : null
              ))
            }


          </div>
        ))
      ) : (
        <p>No values available.</p>
      )}

      {filterCommentData.map((items, index) => (
        <div key={index}>
          <p>{items.Content} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;{items.CreationDate.toLocaleString()}</p>
        </div>
      ))}
      
      
      {
        updateComment.length > 0 &&
        <div>
            <p>{updateComment[0]?.comment} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;{new Date().toLocaleString()} </p>
            </div>
      }
       
          
     
    
    </div>
  );
}

export default TextInsertion;
