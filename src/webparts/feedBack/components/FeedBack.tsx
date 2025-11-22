import * as React from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";

import { IFeedBackProps } from "./IFeedBackProps";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import styles from "./FeedBack.module.scss";

interface IState {
  name: string;
  title: string;
  email: string;
  feedback: string;
  loading: boolean;
  success: boolean;
  error: string | null;
  data: any[];
}

export default class FeedBack extends React.Component<IFeedBackProps, IState> {
  constructor(props: IFeedBackProps) {
    super(props);

    this.state = {
      name: "",
      title: "",
      email: "",
      feedback: "",
      loading: false,
      success: false,
      error: null,
      data: [],
    };
  }

  // -----------------------------------------
  //  VALIDATION FUNCTION
  // -----------------------------------------
  private validateForm(): boolean {
    const { title, name, email, feedback } = this.state;

    if (!title.trim()) {
      this.setState({ error: "Title is required." });
      return false;
    }
    if (!name.trim()) {
      this.setState({ error: "Name is required." });
      return false;
    }
    if (!email.trim()) {
      this.setState({ error: "Email is required." });
      return false;
    }

    // Simple email format validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      this.setState({ error: "Please enter a valid email address." });
      return false;
    }

    if (!feedback.trim()) {
      this.setState({ error: "Feedback cannot be empty." });
      return false;
    }

    return true; // All validations passed
  }

  // -----------------------------------------
  //  SUBMIT FUNCTION
  // -----------------------------------------
  private async submitFeedback(): Promise<void> {
    // Run validations before submit
    if (!this.validateForm()) return;

    this.setState({ loading: true, success: false, error: null });

    try {
      const list = this.props.sp.web.lists.getByTitle("Feedback");

      await list.items.add({
        Title: this.state.title,
        Name: this.state.name,
        EmaiId: this.state.email, // change internal name as required
        FeedBack: this.state.feedback,
      });
    } catch (err) {
      console.error(err);
      this.setState({ error: "Error submitting feedback." });
      return;
    } finally {
      this.setState({ loading: false });
    }

    // Reset on success
    this.setState({
      success: true,
      name: "",
      title: "",
      email: "",
      feedback: "",
    });
  }

  //Get The Data Function

  private async getData(): Promise<void> {
    try {
      const list = this.props.sp.web.lists.getByTitle("FeedBack");
      const items = await list.items
        .select("ID", "Title", "Name", "EmaiId", "FeedBack")
        .top(5)();
      console.log("Items", items);
      this.setState({
        data: items,
      });
    } catch (err) {
      console.error(err);
    }
  }

  public async componentDidMount(): Promise<void> {
    await this.getData();
  }

  //Edite Items Function
  private async onEdit(id: number): Promise<void> {
    // alert("Edit ID: " + id);
    try {
      const list = this.props.sp.web.lists.getByTitle("FeedBack");
      const item = await list.items
        .getById(id)
        .select("Title", "Name", "EmaiId", "FeedBack")();
      console.log("Edit Item:", item);
      this.setState({
        title: item.Title,
        name: item.Name,
        email: item.EmaiId,
        feedback: item.FeedBack,
      });
    } catch (error) {
      console.error("Edit Error:", error);
    }
  }
  //Delete Items Function
  private async onDelete(id: number): Promise<void> {
    alert("Delete ID: " + id);
    try {
      const lists = this.props.sp.web.lists.getByTitle("FeedBack");
      await lists.items.getById(id).delete();
      await this.getData(); // Refresh the data after deletion
    } catch (error) {
      console.error("Delete Error:", error);
    }
  }
  // -----------------------------------------
  //  UI RENDER
  // -----------------------------------------
  public render(): React.ReactElement {
    console.log("Feedback Data:", this.state.data);
    return (
      <div className={styles.card}>
        <h3 className={styles.heading}>Feedback Form</h3>

        {this.state.success && (
          <MessageBar messageBarType={MessageBarType.success}>
            Submitted Successfully!
          </MessageBar>
        )}

        {this.state.error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {this.state.error}
          </MessageBar>
        )}

        <TextField
          label="Title"
          value={this.state.title}
          onChange={(e, v) => this.setState({ title: v || "" })}
          required
        />

        <TextField
          label="Name"
          value={this.state.name}
          onChange={(e, v) => this.setState({ name: v || "" })}
          required
        />

        <TextField
          label="Email"
          value={this.state.email}
          onChange={(e, v) => this.setState({ email: v || "" })}
          required
        />

        <TextField
          label="Feedback"
          multiline
          rows={4}
          value={this.state.feedback}
          onChange={(e, v) => this.setState({ feedback: v || "" })}
          required
        />

        <div style={{ marginTop: 12 }}>
          <PrimaryButton
            text="Submit"
            onClick={() => this.submitFeedback()}
            disabled={this.state.loading}
          />

          <DefaultButton
            text="Reset"
            style={{ marginLeft: 8 }}
            onClick={() =>
              this.setState({
                title: "",
                name: "",
                email: "",
                feedback: "",
                success: false,
                error: null,
              })
            }
          />
        </div>

        {this.state.loading && (
          <Spinner size={SpinnerSize.medium} label="Saving..." />
        )}
        <div>
          <h2 className={styles.heading}>Feedback Records</h2>

          <div className={`${styles["table-wrapper"]}`}>
            <table className={`${styles["responsive-table"]}`}>
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Name</th>
                  <th>Email</th>
                  <th>Feedback</th>
                  <th>Actions</th>
                </tr>
              </thead>

              <tbody>
                {this.state.data.length === 0 ? (
                  <tr>
                    <td
                      colSpan={5}
                      style={{ textAlign: "center", padding: "12px" }}
                    >
                      No Records Found
                    </td>
                  </tr>
                ) : (
                  this.state.data.map((item) => (
                    <tr key={item.ID}>
                      <td data-label="Title">{item.Title}</td>
                      <td data-label="Name">{item.Name}</td>
                      <td data-label="Email">{item.EmaiId}</td>
                      <td data-label="Feedback">{item.FeedBack}</td>

                      {/* ACTION BUTTONS */}
                      <td data-label="Actions">
                        <button
                          style={{
                            marginRight: "8px",
                            padding: "6px 12px",
                            borderRadius: "5px",
                            border: "none",
                            cursor: "pointer",
                            backgroundColor: "#0078d4",
                            color: "white",
                          }}
                          onClick={() => this.onEdit(item.ID)}
                        >
                          Edit
                        </button>

                        {/* DELETE BUTTON (optional) */}

                        <button
                          style={{
                            padding: "6px 12px",
                            borderRadius: "5px",
                            border: "none",
                            cursor: "pointer",
                            backgroundColor: "#d9534f",
                            color: "white",
                          }}
                          onClick={() => this.onDelete(item.ID)}
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }
}
