import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './UserInfoWebPart.module.scss';

import * as pnp from 'sp-pnp-js';

export interface IUserInfoWebPartProps {
}

let property_array = [
  //"AccountName",
  "FirstName",
  "LastName",
  "Title",
  "SPS-UserPrincipalName"];

export default class UserInfoWebPart extends BaseClientSideWebPart<IUserInfoWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit();
  }

  
  private getUserProperties(): void { 

    //FIX
    pnp.setup({
      sp: {baseUrl: window.location.protocol + "//" + window.location.hostname }
    });

    pnp.sp.profiles.myProperties.get().then(function(result) {  
        var userProperties = result.UserProfileProperties;  
        var userPropertyValues = "";  

        userProperties.forEach(function(property) {  
            if (property_array.indexOf(property.Key) > -1  || property_array.indexOf("*") > -1 ) {
              userPropertyValues += "<li><b>" + property.Key + "</b>" + ": " + property.Value + "</li>"; 
            } 
        });  

        document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;  
    }).catch(function(error) {  
        console.log("Error: " + error);  
    });  
}  


  public render(): void {
    this.domElement.innerHTML = `<div class="${ styles.userInfo }">
    <h2 class="${ styles.header_presentazione }">Benvenuto/a,<br>Ecco le tue info</h2>
      <div class="${ styles.userInfo }">
      <ul id="spUserProfileProperties" class="${ styles.UserProfilePropertiesList }"/>
      </ul>
      </div>  
    </div>`;

    this.getUserProperties();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
