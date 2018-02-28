import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { ISPListData } from '@microsoft/sp-page-context/lib/SPList';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { assign, autobind } from 'office-ui-fabric-react/lib/Utilities';
import {
  SPHttpClient, SPHttpClientBatch, SPHttpClientResponse, SPHttpClientConfiguration
} from '@microsoft/sp-http';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { Promise } from 'es6-promise';
import * as lodash from 'lodash';
import * as jquery from 'jquery';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { RadioGroup, Radio } from 'react-radio-group'


export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, {}> {

  public state: IEmployeeDirectoryProps;
  constructor(props, context) {
    super(props);
    this.state = {
      description: "",
      siteurl: this.props.siteurl,
      searchQuerys: "",
      UserName: "",
      UserContact: "",
      UserArray: [],
      options: "",
      loading: false,
      Nodata: false,
      selectedValue: "Employees",
    };
    this.handleChange = this.handleChange.bind(this);
    this.search = this.search.bind(this);
  }

  search(val) {
  }

  private handleChange(value) {
    this.setState({ selectedValue: value, UserArray: [] });
  }

  onKeyPress(event) {
    if (event.charCode === 13) {
      event.preventDefault();
    }
    if (event.key === 'Enter') {
      event.preventDefault();
    }
  }

  public OnchangeBuilding(event: any): void {
    this.setState({ searchQuerys: event.target.value });
    console.log(this.state.searchQuerys);
  }


  componentDidMount() {
    this.GetUSerDetails();
  }

  private GetUSerDetails() {
    var reqUrl = "https://**************/_api/sp.userprofiles.peoplemanager/GetMyProperties";
    jquery.ajax(
      {
        url: reqUrl, type: "GET", headers:
          {
            "accept": "application/json;odata=verbose"
          }
      }).then((response) => {
        console.log(response.d);
        var Name = response.d.DisplayName;
        var email = response.d.Email;
        var oneUrl = response.d.PersonalUrl;
        var imgUrl = response.d.PictureUrl;
        var jobTitle = response.d.Title;
        var profUrl = response.d.UserUrl;
        var Department = response.d["UserProfileProperties"].results[11];
        this.setState({ searchQuerys: Department.Value, querytext: Department.Value });
        this.RequestSearch();
      });
  }

  private RequestSearch(): void {
    var reactHandler = this;
    reactHandler.setState({ Nodata: false });
    reactHandler.setState({ loading: true });
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("SitePages", "");
    console.log(NewSiteUrl);
    var XSource = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";
    var SelectProp = "&selectproperties='AccountName,WorkPhone,WorkEmail,Department,AccountName,FirstName,LastName,AboutMe,PictureUrl,JobTitle'";
    var Searchurl = "https://Enter your tenant url/_api/search/query?querytext='" + this.state.searchQuerys + "*'&sourceid='" + XSource + "'" + SelectProp + "&ROWLIMIT=1000";
    if (this.state.selectedValue != "Employees") {
      var Searchurl = "https://Enter your tenant url/_api/search/query?querytext='" + this.state.searchQuerys + "*'&selectproperties='Department,WorkPhone,WorkEmail'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&rowlimit=1000";
    }
    console.log(Searchurl);
    jquery.ajax({
      url: `${Searchurl}`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        var myObject = JSON.stringify(resultData.d.results);
        var results = resultData.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
        reactHandler.setState({ UserArray: results, loading: false });

        if (results.length == 0) {
          reactHandler.setState({ Nodata: true });
        }
        else {
          reactHandler.setState({ Nodata: false });
          var searchResultsHtml = '';
          jquery.each(results, function (index, result) {
            searchResultsHtml += "<a target='_blank' href='" + result.Cells.results[6].Value + "'>" + result.Cells.results[3].Value + "</a> (" + result.Cells.results[10].Value + ")<br />";
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }




  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    let NodataConent;
    if (this.state.Nodata == true) {
      NodataConent = <div className={styles.NoResultDiv} ><h3>No Result Found.</h3></div>
    } else {
      NodataConent = <div></div>;
    }

    let content;
    if (this.state.loading) {
      content = <div><img src="https://********************/sites/dev/SiteAssets/loadingnew.gif" /></div>;
    } else {
      content = <div></div>;
    }
    var renderItems;
    if (this.state.selectedValue == "Employees") {
      renderItems = this.state.UserArray.map(function (item, i) {
        if (item["Cells"]["results"][4].Value != null) {
          if (item["Cells"]["results"][9].Value == null) {
            item["Cells"]["results"][9].Value = "https://******************/sites/spdemo/SiteAssets/emptyphoto.png";
          }
          let MyNewImageValue = encodeURI(item["Cells"]["results"][9].Value);
          return <div className={styles.row}>
            <div className={styles.mydivs} >
              <img className={styles.myimage} src={MyNewImageValue} />
              <h5>
                {item["Cells"]["results"][6].Value}  {item["Cells"]["results"][7].Value}
              </h5>
            </div>
            <div className={styles.mydivs2} >
              <h5>
                <img src="https://cdn1.iconfinder.com/data/icons/office-and-employment-vol-2/32/199_office_employee_bedge_label_tag_award_ribbon_number_one_first_position_best-128.png" className={styles.mySmallImagesIcon} />
                {item["Cells"]["results"][10].Value}</h5>
              <h5>
                <img src="https://image.flaticon.com/icons/png/128/254/254048.png" className={styles.mySmallImagesIcon} />
                <a href={item["Cells"]["results"][4].Value}>{item["Cells"]["results"][4].Value}</a>
              </h5>
              <h5>
                <img src="https://n6-img-fp.akamaized.net/free-icon/3d-buildings_318-79332.jpg" className={styles.mySmallImagesIcon} />
                {item["Cells"]["results"][5].Value}  </h5>
              <h5>
                <img src="http://clipart-library.com/img/1344166.png" className={styles.mySmallImagesIcon} />
                {item["Cells"]["results"][3].Value}  </h5>
            </div>

          </div>
        }

      });
    } if (this.state.selectedValue == "Department") {
      let MyArray = [];
      renderItems = this.state.UserArray.map(function (item, i) {
        if (MyArray.indexOf(item["Cells"]["results"][2].Value) < 0 && item["Cells"]["results"][2].Value != null) {
          //start
          MyArray.push(item["Cells"]["results"][2].Value);
          return <div className={styles.row}>
            <div className={styles.mydivs2} >
              <h5>
                <img src="https://n6-img-fp.akamaized.net/free-icon/3d-buildings_318-79332.jpg" className={styles.mySmallImagesIcon} />
                {item["Cells"]["results"][2].Value}  </h5>
              <h5>
                <img src="http://clipart-library.com/img/1344166.png" className={styles.mySmallImagesIcon} />
                {item["Cells"]["results"][3].Value}  </h5>
            </div>
          </div>
          //END
        }
      });
    }
    return (
      <div className={styles.employeeDirectory}>
        <div className={styles.row}><h2>Employee Directory</h2>{this.state.searchQuerys}</div>
        <div className="ms-TextField">
          <input type="text" ref="myinputtext" className={styles.myinput} onKeyPress={this.onKeyPress.bind(this)} onChange={this.OnchangeBuilding.bind(this)} placeholder="Enter your Text here..." />
          <div className={styles.row}>
            <RadioGroup name="fruit" selectedValue={this.state.selectedValue} onChange={this.handleChange}>
              <Radio value="Employees" className={styles.myinputRadio} />Employees
                <Radio value="Department" className={styles.myinputRadio} />Department
              </RadioGroup>
          </div>
        </div>

        <div className={styles.row}>
          <div  >
            <button autoFocus={true} type="button" id="btn_add" className={styles.button} onClick={this.RequestSearch.bind(this)}>Search </button>
          </div>
        </div>
        {this.state.options}
        <div className={styles.results}>
          {NodataConent}
          {content}
          {renderItems}
        </div>

      </div>
    );


  }


}
