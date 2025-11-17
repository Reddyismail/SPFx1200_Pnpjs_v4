import * as React from "react";
// import styles from './TeamsTab.module.scss';
import type { ITeamsTabProps } from "./ITeamsTabProps";
// import { escape } from '@microsoft/sp-lodash-subset';
// import * as microsoftTeams from '@microsoft/teams-js';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";

interface IState {
  items: any[];
}
export default class TeamsTab extends React.Component<ITeamsTabProps, IState> {
  constructor(props: ITeamsTabProps) {
    super(props);
    this.state = { items: [] };
  }

  public async componentDidMount() {
    await this.loadItems();
  }
  private async loadItems() {
    // Get all items (default returns 100 items max)
    // const items = await this.props.sp.web.lists
    //   .getByTitle("CustomerData_3k") // your SharePoint list name
    //   .rootFolder.folders(); // execute query
    const items = await this.props.sp.web.folders(); // execute query

    console.log("folders:", items);
    this.setState({ items }); // store in React state
  }

  public render(): React.ReactElement<ITeamsTabProps> {
    return (
      <div style={{ padding: "20px", fontFamily: "Segoe UI" }}>
        <h2 style={{ marginBottom: "20px" }}>Library Folders</h2>

        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            border: "1px solid #ddd",
          }}
        >
          <thead>
            <tr>
              <th
                style={{
                  padding: "8px",
                  border: "1px solid #ddd",
                  background: "#f3f3f3",
                }}
              >
                Folder Name
              </th>
              <th
                style={{
                  padding: "8px",
                  border: "1px solid #ddd",
                  background: "#f3f3f3",
                }}
              >
                Server Relative Url
              </th>
            </tr>
          </thead>

          <tbody>
            {this.state.items.map((f) => (
              <tr key={f.Name}>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                  {f.Name}
                </td>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                  {f.ServerRelativeUrl}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}
