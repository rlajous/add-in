import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      value: "",
    };

    this.handleChange = this.handleChange.bind(this);
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(function (event) {
        return Excel.run(function (context) {
          this.setState({
            value: event.toString(),
          });
          console.log("The selected range has changed to: " + event.address);
          return context.sync();
        });
      });

      console.log("The worksheet click handler is registered.");

      await context.sync();
    });
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSelectionChanged.add(function (event) {
        return Excel.run(function (context) {
          console.log("The selected range has changed to: " + event.address);
          return context.sync();
        });
      });
    }).catch();
  }

  handleChange(event) {
    this.setState({ value: event.target.value });
  }

  handleChangeExcel(event) {
    return Excel.run(function (context) {
      return context.sync().then(function () {
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
        this.setState({
          value: event,
        });
      });
    }).catch();
  }

  click = async () => {
    const reader = new FileReader();

    try {
      await Excel.run(async (context) => {
        const startIndex = reader.result.toString().indexOf("base64,");

        context;

        // 7 is the length of the "base64," string to skip past
        const workbookContents = reader.result.toString().substr(startIndex + 7);

        // STEP 1: Insert the template into the workbook.
        const workbook = context.workbook;

        // Set up the insert options.
        var options = {
          sheetNamesToInsert: ["Template"], // Insert the "Template" worksheet from the source workbook.
          positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
          relativeTo: "Sheet1",
        }; // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.

        // Insert the external worksheet.
        workbook.insertWorksheetsFromBase64(workbookContents, options);

        // In Excel on the web, if the worksheet being inserted contains unsupported features,
        // such as Comment, Slicer, Chart, and PivotTable, insertWorksheetsFromBase64 will fail.
        // In your production add-in, you should notify the user in the add-ins UI.
        // As a workaround they can use Excel on desktop, or choose a different worksheet.
        await context.sync();
        /**
         * Insert your Excel code here
         */
        //const range = context.workbook.getSelectedRange();

        // Read the range address
        //range.load("address");

        // Update the fill color
        //range.format.fill.color = "yellow";

        /*

        var sheets = context.workbook.worksheets;

        var sheet = sheets.add("So");
        sheet.load("name, position");
    
        return context.sync()
            .then(function () {
              console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
            });
            

        await context.sync();
        console.log(`The range address was ${range.address}.`);
        */
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <input type="text" value={this.state.value} onChange={this.handleChange} />
          {this.state.value}
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.handleChangeExcel}
          >
            Runa
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
