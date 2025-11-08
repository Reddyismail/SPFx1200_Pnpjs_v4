import * as React from "react";
import styles from "./CallApi.module.scss";
import type { ICallApiProps } from "./ICallApiProps";
// import { escape } from "@microsoft/sp-lodash-subset";
import { getRandomString } from "@pnp/core";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/search";
// import { ISearchQuery, SearchResults, SearchQueryBuilder } from "@pnp/sp/search";

interface IState {
  items: any[];
  Serach: any[];
}
export default class CallApi extends React.Component<ICallApiProps, IState> {
  constructor(props: ICallApiProps) {
    super(props);
    this.state = { items: [], Serach: [] };
  }
  public async componentDidMount() {
    await this.loadItems();
    await this.loadSerach();
  }
  private async loadItems() {
    // Get all items (default returns 100 items max)
    const items = await this.props.sp.web.lists
      .getByTitle("CustomerData_3k") // your SharePoint list name
      .items.select("ID", "Title", "field_1", "field_2", "field_3")
      .top(100)(); // execute query

    console.log("Items loaded:", items);
    this.setState({ items }); // store in React state
  }
  private async loadSerach() {
    // Get all items (default returns 100 items max)
    const ite = await this.props.sp.search("test");
    console.log("Items loaded:", ite);
    //this.setState({ Serach }); // store in React state
  }
  public render(): React.ReactElement<ICallApiProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName,
    // } = this.props;

    console.log("Random String: ", getRandomString(20));
    return (
      <div>
        <div className={`${styles["table-wrapper"]}`}>
          <table className={`${styles["responsive-table"]}`}>
            <thead>
              <tr>
                <th>Name</th>
                <th>Full Name</th>
                <th>Age</th>
                <th>Contact</th>
              </tr>
            </thead>

            <tbody>
              {this.state.items.map((item) => (
                <tr key={item.ID}>
                  <td data-label="Name">{item.Title}</td>
                  <td data-label="Full Name">{item.field_1}</td>
                  <td data-label="Age">{item.field_2}</td>
                  <td data-label="Contact">{item.field_3}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  }
}
