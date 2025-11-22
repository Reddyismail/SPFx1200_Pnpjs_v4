import * as React from "react";
// import styles from "./Permission.module.scss";
import type { IPermissionProps } from "./IPermissionProps";
// import { escape } from "@microsoft/sp-lodash-subset";
// import {sp} from "@pnp/sp/presets/all";/
import { spfi,SPFx,SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
// import {spfi} from "@pnp/sp/presets/all";


export interface IState {
  canEdit: boolean;
  isManager: boolean;
}

export default class Permission extends React.Component<IPermissionProps,IState> {
private sp: SPFI;
constructor(props: any) {
    super(props);

    this.state = {
      canEdit: false,
      isManager: false
    };
    this.sp = spfi().using(SPFx(this.props.context));
  // public async componentDidMount():Promise<void> {
  // await this.checkPermissions();
   }
  
  public async componentDidMount():Promise<void> {
    await this.checkPermissions();
  }
  private async checkPermissions():Promise<void>  {

    // ⭐ 1. Check EditListItems permission → PnP v4 uses numeric values
    // const canEdit = await this.sp.web.currentUserHasPermissions(2048); // EditListItems

    // ⭐ 2. Check if user belongs to "Managers" group
    const groups = await this.sp.web.currentUser.groups();
   const isOwner = groups.some(g => g.Title.toLowerCase().includes("owners"));
const isMember = groups.some(g => g.Title.toLowerCase().includes("members"));
const isVisitor = groups.some(g => g.Title.toLowerCase().includes("visitors"));


console.log("Is Owner:", isOwner);
console.log("Is Member:", isMember);
console.log("Is Visitor:", isVisitor);  
let permission = "";
console.log(groups)
if (isOwner) permission = "Full Control";
else if (isMember) permission = "Edit";
else if (isVisitor) permission = "Read";
else permission = "No Access";

console.log(permission);

    this.setState({
      // canEdit,
      isManager:isMember
    });
  }
//   private async checkPermissions():Promise<void>{
// const canEdit=await sp.web.currentUserHasPermissions(sp.PermissionKind.EditListItems);
// console.log("User can edit items:",canEdit);
//   }
  public render(): React.ReactElement<IPermissionProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName,
    // } = this.props;
    console.log("Image URL:", this.props.image);

    // const iconUrl = `${this.props.absoluteUrl}/${this.props.image}`;

    return <div>Hi</div>;
  }
}
