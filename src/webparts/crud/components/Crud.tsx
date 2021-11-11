import * as React from "react";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/profiles";

import {
  ColorPicker,
  DatePicker,
  DefaultButton,
  DetailsList,
  getColorFromString,
  IColumn,
  IPersonaProps,
  NormalPeoplePicker,
  Panel,
  PrimaryButton,
  Stack,
  TextField,
} from "office-ui-fabric-react";

import styles from "./Crud.module.scss";
import { ICrudProps, ICrudState } from "./Crud.types";

export default class Crud extends React.Component<ICrudProps, ICrudState> {
  constructor(props: ICrudProps) {
    super(props);

    const columns: IColumn[] = [
      {
        key: "column0",
        name: "Id",
        fieldName: "ID",
        minWidth: 25,
        maxWidth: 50,
        isResizable: true,
      },
      {
        key: "column1",
        name: "Associate Name",
        fieldName: "AssociateName",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: this._onColumnClick,
      },
      {
        key: "column2",
        name: "Age",
        fieldName: "Age",
        minWidth: 25,
        maxWidth: 50,
        isResizable: true,
        isSorted: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: this._onColumnClick,
      },
      {
        key: "column3",
        name: "Date Of Joining",
        fieldName: "Date",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: true,
        sortAscendingAriaLabel: "Newer to Older",
        sortDescendingAriaLabel: "Older to Newer",
        onColumnClick: this._onColumnClick,
      },
    ];

    this.state = {
      associateName: "",
      age: "",
      date: new Date(),
      allAssociates: [],
      open: false,
      selectedAssociate: {},
      color: getColorFromString("#fff"),
      columns,
    };

    sp.setup({
      spfxContext: this.props.spcontext,
    });

    this.fetchData = this.fetchData.bind(this);
    this.submitHandler = this.submitHandler.bind(this);
    this.updateHandler = this.updateHandler.bind(this);
    this.handleDelete = this.handleDelete.bind(this);
  }

  private async fetchData() {
    const items = await sp.web.lists.getByTitle("CRUD").items.getAll();

    const allAssociates = items.map((item) => {
      return {
        ...item,
        Date: new Date(item.Date).toString(),
      };
    });

    this.setState({
      allAssociates: allAssociates,
    });
  }

  public async componentDidMount() {
    await this.fetchData();
    
    const profile = await sp.profiles.userProfile;
    console.log(profile);
  }

  // Form Submit Handler
  private async submitHandler(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();

    if (!this.state.associateName || !this.state.age || !this.state.date) {
      alert("Please fill all the fields");
      return;
    }

    await sp.web.lists.getByTitle("CRUD").items.add({
      AssociateName: this.state.associateName,
      Age: this.state.age,
      Date: this.state.date.toISOString(),
    });

    await this.fetchData();

    this.setState({
      associateName: "",
      age: "0",
      date: new Date(),
    });
  }

  // Delete Handler
  private async handleDelete() {
    await sp.web.lists
      .getByTitle("CRUD")
      .items.getById(this.state.selectedAssociate.ID)
      .delete();

    await this.fetchData();

    this.setState({
      open: false,
      associateName: "",
      age: "",
      date: new Date(),
    });
  }

  // Update Handler
  private async updateHandler() {
    await sp.web.lists
      .getByTitle("CRUD")
      .items.getById(this.state.selectedAssociate.ID)
      .update({
        AssociateName: this.state.associateName,
        Age: this.state.age,
        Date: this.state.date,
      });

    await this.fetchData();

    this.setState({
      open: false,
      associateName: "",
      age: "",
      date: new Date(),
    });
  }

  private _onColumnClick = (e, column: IColumn) => {
    const newColumns = this.state.columns.slice();
    const currColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];

    newColumns.forEach((newCol) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const sortedItems = this.state.allAssociates.sort((a, b) => {
      const isSortedDesc = currColumn.isSortedDescending;

      if (a[currColumn.fieldName] < b[currColumn.fieldName]) {
        return isSortedDesc ? 1 : -1;
      } else if (a[currColumn.fieldName] > b[currColumn.fieldName]) {
        return isSortedDesc ? -1 : 1;
      } else {
        return 0;
      }
    });

    this.setState({
      columns: newColumns,
      allAssociates: sortedItems,
    });
  };

  private onFilterChanged = (
    filterText: string,
    selectedItems?: IPersonaProps[]
  ) => {
    return selectedItems;
  };

  public render(): React.ReactElement<ICrudProps> {
    const { associateName, age, date, open, allAssociates, color } = this.state;

    return (
      <Stack verticalAlign="center">
        <form
          style={{ backgroundColor: `#${color.str}` }}
          onSubmit={this.submitHandler}
        >
          <TextField
            label="Associate"
            value={associateName}
            placeholder="Associate Name"
            onChange={(e, str) => this.setState({ associateName: str })}
            required
          />

          <TextField
            label="Age"
            value={age}
            onChange={(e, str) => this.setState({ age: str })}
            placeholder="Age"
            type="number"
            required
          />

          <DatePicker
            label="DOJ"
            isRequired
            placeholder="Select a date..."
            ariaLabel="Select a date"
            value={date}
            onSelectDate={(date) => this.setState({ date })}
          />

          <Stack styles={{ root: { marginTop: "0.5em" } }}>
            <p style={{ marginTop: "0.25em", marginBottom: "0.25em" }}>
              Reports To:
            </p>
            <NormalPeoplePicker
              className={"ms-PeoplePicker"}
              onResolveSuggestions={this.onFilterChanged}
            />
          </Stack>

          <PrimaryButton style={{ marginTop: "1em" }} type="submit">
            Submit
          </PrimaryButton>
        </form>

        <Panel
          isOpen={open}
          headerText="Edit Associate"
          closeButtonAriaLabel="cancel"
          isFooterAtBottom={true}
          onDismiss={() => this.setState({ open: false })}
          onRenderFooterContent={() => (
            <Stack
              horizontal
              horizontalAlign="center"
              verticalAlign="center"
              gap="1em"
            >
              <PrimaryButton
                onClick={() => {
                  this.setState({ open: false });
                  this.updateHandler();
                }}
                text="Update"
              />
              <DefaultButton onClick={this.handleDelete} text="Delete" />
              <DefaultButton
                onClick={() => this.setState({ open: false })}
                text="Cancel"
              />
            </Stack>
          )}
        >
          <TextField
            label="Associate"
            value={associateName}
            placeholder="Associate Name"
            onChange={(e, str) => this.setState({ associateName: str })}
            required
          />

          <TextField
            label="Age"
            value={age}
            onChange={(e, str) => this.setState({ age: str })}
            placeholder="Age"
            type="number"
            required
          />

          <DatePicker
            label="DOJ"
            isRequired
            placeholder="Select a date..."
            ariaLabel="Select a date"
            value={date}
            onSelectDate={(date) => this.setState({ date })}
          />
        </Panel>

        {allAssociates.length ? (
          // <Stack horizontal gap="1em">
          //  <ColorPicker
          //   color={color}
          //   onChange={(e, newColor) => this.setState({ color: newColor })}
          // />
          // </Stack>

          <DetailsList
            items={allAssociates}
            columns={this.state.columns}
            onItemInvoked={(item) => {
              this.setState({
                open: true,
                selectedAssociate: item,
                associateName: item.AssociateName,
                age: item.Age,
                date: new Date(item.Date),
              });
            }}
          />
        ) : (
          <h3> No Data Found </h3>
        )}
      </Stack>
    );
  }
}
