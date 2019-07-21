import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestFormWebPart.module.scss';
import * as strings from 'TestFormWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult, Web } from 'sp-pnp-js';

import * as $ from 'jquery';
require('bootstrap');
require('./css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
require('appjs');
require('sppeoplepicker');
require('jqueryui');

var queryParms = new UrlQueryParameterCollection(window.location.href);
var SpId = queryParms.getValue("ID");

export interface ITestFormWebPartProps {
  description: string;
}

export default class TestFormWebPart extends BaseClientSideWebPart<ITestFormWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div id="container">
     
    <div class="panel">
      <div class="panel-body">
        <div class="row">
          <div class="col-lg-4 control-padding">
            <label>Activity</label>
            <input type='textbox' name='txtActivity' id='txtActivity' class="form-control" value="" placeholder="" >
          </div>
          <div class="col-lg-4 control-padding">
            <label>Activity Performed By</label>
           
              <div id="ppDefault"></div>
          </div>
          <div class="col-lg-4 control-padding">
            <label>Activity Date</label>
            <div class="input-group date" data-provide="datepicker">
      <input type="text" class="form-control" id="txtDate" name="txtDate">
  </div>
          </div>          
        </div>
  
        <div class="row">
        <div class="col-lg-6 control-padding">
            <label>Category</label>
            <select name="ddlCategory" id="ddlCategory" class="form-control">
  
            </select>
          </div>
          <div class="col-lg-6 control-padding">
            <label>Sub Category</label>
            <select name="ddlSubCategory" id="ddlSubCategory" class="form-control">
  
            </select>
          </div>         
        </div>       
  
        <div class="row">
        <div class="col col-lg-12">
        <button type="button" class="btn btn-primary buttons" id="btnSubmit">Save</button>
			  <button type="button" class="btn btn-default buttons" id="btnCancel">Cancel</button>
      </div>
        </div>
  
      </div>
    </div>`;

    (<any>$("#txtDate")).datepicker(
      {
        changeMonth: true,
        changeYear: true,
        dateFormat: "mm/dd/yy"
      }
    );
    (<any>$('#ppDefault')).spPeoplePicker({
      minSearchTriggerLength: 2,
      maximumEntitySuggestions: 10,
      principalType: 1,
      principalSource: 15,
      searchPrefix: '',
      searchSuffix: '',
      displayResultCount: 6,
      maxSelectedUsers: 1
    });
    this.AddEventListeners();
    this.getCategoryData();
  }

  private AddEventListeners(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this.SubmitData());
    document.getElementById('btnCancel').addEventListener('click', () => this.CancelForm());   
    document.getElementById('ddlCategory').addEventListener('change', () => this.PopulateSubCategory());
  }

private SubmitData(){
  var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
  var userId;
  var userDetails = this.GetUserId(userinfo[0].email.toString());
  console.log(JSON.stringify(userDetails));
  userId = userDetails.d.Id;

  pnp.sp.web.lists.getByTitle('RigActiveList_Job_Cards_Area').items.add({
    Title: "Test",
    Activity: $("#txtActivity").val().toString(),
    Activity_Date: $("#txtDate").val().toString(),
    Activity_ById : userId,
    Category: $("#ddlCategory").val().toString(),
    SubCategory: $("#ddlSubCategory").val().toString(),
});
}

private GetUserId(userName) {
  var siteUrl = this.context.pageContext.web.absoluteUrl;

  var call = $.ajax({
    url: siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",
    method: "GET",
    headers: { "Accept": "application/json; odata=verbose" },
    async: false,
    dataType: 'json'
  }).responseJSON;
  return call;
}

  private CancelForm() {
    window.location.href = this.GetQueryStringByParameter("Source");
  }

  private GetQueryStringByParameter(name) {
    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
      results = regex.exec(location.search);
    return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
  }


  private _getCategoryData(): any {    
    return pnp.sp.web.lists.getByTitle("Category").items.select("Category").getAll().then((response) => {
      return response;
    });
  }

  private getCategoryData(): any {
    this._getCategoryData()
      .then((response) => {
        this._renderCategoryList(response);
      });
  }

  private _renderCategoryList(items: any): void {

    let html: string = '';
    html += `<option value="Select Category" selected>Select Category</option>`;
    items.forEach((item: any) => {
      html += `
       <option value="${item.Category}">${item.Category}</option>`;
    });
    const listContainer1: Element = this.domElement.querySelector('#ddlCategory');
    listContainer1.innerHTML = html;
  }

  public PopulateSubCategory() {
    this.getSubCategoryData($("#ddlCategory").val().toString());
  }

  private _getSubCategoryData(category): any {    
    return pnp.sp.web.lists.getByTitle("SubCategory").items.select("SubCategory").filter("Category eq '" + category + "'").getAll().then((response) => {
      return response;
    });
  }

  private getSubCategoryData(category): any {
    this._getSubCategoryData(category)
      .then((response) => {
        this._renderSubCategoryList(response);
      });
  }

  private _renderSubCategoryList(items: any): void {

    let html: string = '';
    html += `<option value="Select Sub Category" selected>Select Sub Category</option>`;
    items.forEach((item: any) => {
      html += `
       <option value="${item.SubCategory}">${item.SubCategory}</option>`;
    });
    const listContainer1: Element = this.domElement.querySelector('#ddlSubCategory');
    listContainer1.innerHTML = html;
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
