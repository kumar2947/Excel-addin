import React, { useState, useEffect } from "react";
import axios from 'axios';

function TextInsertion() {
  const [updateComment, setUpdatedComment] = useState([]);
  const [allComments, setAllComments] = useState([]);
  const [excelData, setExcelData] = useState([])

  useEffect(() => {
    // Make a GET request to the posts endpoint using axios
    axios.get('http://localhost:3002/comments')
      .then(response => setAllComments(response.data))
      .catch(error => console.error('Error fetching data:', error));
  }, []);

  useEffect(() => {
    // Make a GET request to the posts endpoint using axios
  
  }, []);
const [updateC, setUpdateC] = useState(false)
  useEffect(() => {
  
    initializeEventHandlers();
  }, []);



  async function initializeEventHandlers() {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem("Sheet1");
        worksheet.onChanged.add(handleChange);

        await context.sync();
        console.log("Event handler successfully registered for onChanged event in the worksheet.");
      });
    } catch (error) {
      console.error("Error initializing event handlers: " + error.message);
    }

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        let range = worksheet.getRange();
        range.clear();
     
        // Create the headers and format them to stand out.
        let headers = [
          ["ID", "MOBILENO"]
        ];
        let headerRange = worksheet.getRange("A1:B1");
        headerRange.values = headers;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
        headerRange.format.columnWidth = 130;
     
        axios.get('http://localhost:3002/spPersonalData')
        .then(response => {
          // Handle the data within the 'then' callback
             let excelFormatData = [];
   
        response.data.forEach(element => {
          excelFormatData.push([element.pId,element.mobileNo])
        })

        let dataRange = worksheet.getRange("A2:B4");
        dataRange.values = excelFormatData;
          // You can now use 'excelData' or any logic dependent on the data here

          // Register event handlers
          worksheet.onChanged.add(handleChange);
          worksheet.comments.onChanged.add(commentChanged);
          worksheet.comments.onAdded.add(commentAdded);

          console.log("Event handlers successfully registered.");
        })
        .catch(error => console.error('Error fetching data:', error));
     


 

        await context.sync();
        console.log("Event handlers successfully registered.");
      });
    } catch (error) {
      console.error("Error initializing event handlers: " + error.message);
    }


  }

  async function commentChanged(event) {
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
      const newComment = {
        comment: addedComment.content,
        createdBy: addedComment.authorName,
        created: addedComment.creationDate,
      };
      setUpdatedComment((prevComments) => [
        ...prevComments,
        newComment,
      ]);

      try {
        const response = await axios.post('http://localhost:3002/comments', {
          comment: addedComment.content,
          createdBy: addedComment.authorName,
          created: addedComment.creationDate
        });

        console.log('New post added:', response.data);
        // You may update the UI or perform other actions after a successful request.
      } catch (error) {
        console.error('Error adding post:', error);
      }

      console.log(`A comment was added:`);
      console.log(`    ID: ${event.commentDetails[0].commentId}`);
      console.log(`    Comment content:${addedComment.content}`);
      console.log(`    Comment author:${addedComment.authorName}`);
      console.log(`    Creation date:${addedComment.creationDate}`);
    });
  }

  async function handleChange(event) {
    console.log(event)
    await Excel.run(async (context) => {
      await context.sync();
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
      console.log("Before Value: " + event.details.valueBefore);
      console.log("After Value: " + event.details.valueAfter)
    }).catch(errorHandlerFunction);

  }
  const handleDataChanged = (eventArgs) => {
    if (eventArgs && eventArgs.type === Office.EventType.BindingDataChanged) {
      const bindingData = eventArgs.bindingData;
      if (bindingData && bindingData.type === "comment") {
        setUpdatedComment(bindingData.commentText);
      }
    }
  };

  // const getDataFromExcel = async () => {
  //   try {
  //     const data = [];

  //     await Excel.run(async (context) => {
  //       const sheet = context.workbook.worksheets.getActiveWorksheet();
  //       const range = sheet.getUsedRange();

  //       range.load("values");

  //       await context.sync();

  //       const values = range.values;

  //       for (let i = 0; i < values.length; i++) {
  //         const rowData = {};

  //         for (let j = 0; j < values[i].length; j++) {
  //           const header = values[0][j];
  //           const cellValue = values[i][j];

  //           // Use the header as the property name and cellValue as the value.
  //           rowData[header] = cellValue;
  //         }

  //         data.push(rowData);
  //       }
  //     });

  //     // Skip the first row (header) and return the rest as JSON data.
  //     const jsonData = JSON.stringify(data.slice(1));

  //     // Log the data for demonstration
  //     console.log(jsonData);
  //   } catch (error) {
  //     console.error("Error: " + error);
  //   }
  // };

  console.log(updateComment)
  const spreadData = () => {

    // await Excel.run(async (context) => {
    //   let sheet = context.workbook.worksheets.getItem("Sample");
    //   let expensesTable = sheet.tables.getItem("ExpensesTable");

    //   expensesTable.rows.add(
    //       null, // index, Adds rows to the end of the table.
    //       [
    //           ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
    //           ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
    //           ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
    //           ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
    //           ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
    //           ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
    //           ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
    //       ], 
    //       true, // alwaysInsert, Specifies that the new rows be inserted into the table.
    //   );

    //   sheet.getUsedRange().format.autofitColumns();
    //   sheet.getUsedRange().format.autofitRows();

    //   await context.sync();
    // });
  }

  return (
    <div>
      <h1>Audit History</h1>
      {/* <button onClick={spreadData}>Spread Data</button> */}

      {allComments && allComments.length > 0 ? (

        allComments.map((item, index) => (
          <div key={index}>
            <p>{item.comment}</p>
            <p>Created by: {item.createdBy}</p>
            <p>Created on: {item.created.toLocaleString()}</p> {/* Format the date */}
          </div>
        ))


      ) : (
        <p>No comments available.</p>
      )}


    </div>
  );
}

export default TextInsertion;
