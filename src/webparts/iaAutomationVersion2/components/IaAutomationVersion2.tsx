import * as React from 'react';
import styles from './IaAutomationVersion2.module.scss';
import { IIaAutomationVersion2Props } from './IIaAutomationVersion2Props';
import { escape } from '@microsoft/sp-lodash-subset';

import {ChoiceGroup,IChoiceGroupOption, textAreaProperties,Stack, IStackTokens, StackItem,IStackStyles,TextField } from 'office-ui-fabric-react'; 

import {Icon} from 'office-ui-fabric-react/lib/Icon';

import {DatePicker} from 'office-ui-fabric-react';

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import Service from './Service';

import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';  

//#region GlobalVaraibles
const sectionStackTokens: IStackTokens = { childrenGap: 10 };
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { padding: 10} };
const stackButtonStyles: Partial<IStackStyles> = { root: { width: 20 } };

import { Button,PrimaryButton } from 'office-ui-fabric-react/lib/Button';

const RadioServiceNow: IChoiceGroupOption[] = 

[  { key: "Yes", text: "Yes" , },  { key: "No", text: "No" },];  

const RadioSystem: IChoiceGroupOption[] = 

[  { key: "PeopleSoft FIN", text: "PeopleSoft FIN" , },  { key: "PeopleSoft HR", text: "PeopleSoft HR" },{ key: "Greenhouse", text: "Greenhouse" },{ key: "Rydoo", text: "Rydoo" },{ key: "Concur", text: "Concur" },{ key: "Egencia", text: "Egencia"},{ key: "Cornerstone", text: "Cornerstone"},{ key: "SharePoint", text: "SharePoint"},{ key: "Multiple Systems", text: "Multiple Systems"},{ key: "Power BI", text: "Power BI"} ];  

let RootUrl = '';

let RequestorName='';

let ClientPartnerName='';

let ReqSiteUrl='';

let FinalStatus='';

let GreatUrl='';

let Disablevalue=null;



let mystring='Refrain from using the following special characters in your title: ? * \ / : < > | & # " %';

export interface IAAutomationControlFieldsState2{

  ReqName:any;
  dtreqdate:Date;
  System:any;
  Title:any;
  Desc:any;
  dtReqDelivery:any;
  FileValue:any;
 disableFileUpload:boolean;
 ServiceNow:any;
 ServiceNowTicketNumber:any;
 dtServTickOpen:any;
 Email:any;
 flag: boolean;
 ClientPatnerName:any;
 ServiceFlag:boolean;
 UserLoginName:any;
 userval:any;
 dtserviceNowdate:any;
 MyUrl:any;
 MyListName:any;
 MyColNameDesc:any;
 ServiceNowtext:any;
 FinalUrl:any;
 savedisabled:boolean;

}

export default class IaAutomationVersion2 extends React.Component<IIaAutomationVersion2Props, IAAutomationControlFieldsState2> {
  public _service: any;
  public GlobalService: any;
  protected ppl;

  public constructor(props:IIaAutomationVersion2Props){
    super(props);
    this.state={
      ReqName:null,
      dtreqdate:null,
      System:null,
      Title:null,
      Desc:"",
      dtReqDelivery:null,
      FileValue:[],
      disableFileUpload:false,
      ServiceNow:null,
      ServiceNowTicketNumber:null,
      dtServTickOpen:null,
      Email:null,
      flag: false,
      ServiceFlag:false,
      UserLoginName:"",
      userval:[],
      ClientPatnerName:"",
      dtserviceNowdate:null,
      MyUrl:"",
      MyListName:"",
      MyColNameDesc:"",
      ServiceNowtext:null,
      FinalUrl:"",
      savedisabled:true

    };


    // RootUrl = this.props.url;

    // this._service = new Service(this.props.url, this.props.context);

    // this.GlobalService = new Service(this.props.url, this.props.context);

//

    

    RootUrl = "https://capcoinc.sharepoint.com/sites/capcointernalapplications/";

    this._service = new Service("https://capcoinc.sharepoint.com/sites/capcointernalapplications/", this.props.context);

    this.GlobalService = new Service("https://capcoinc.sharepoint.com/sites/capcointernalapplications/", this.props.context);

     this.getUserDetails();

    this.getListNameandURL();

    
  }

 

  public async getListNameandURL()
  {

    var data = await this._service.GetListNameandURL();

    console.log(data);

    this.setState({MyUrl:data[0].URL,MyListName:data[0].Title,MyColNameDesc:data[0].ColName,FinalUrl:data[0].MainUrl});

  }
 

  private changeEmail(data: any): void {

    this.setState({ Email: data.target.value });

  }

  private changeDesc(data: any): void {

    this.setState({ Desc: data.target.value });

  }

  private changeTitle(data: any): void {

    this.setState({ Title: data.target.value.replace(/[`#%&*|\?:'"<>\\/]/gi, '')});

  }

  private changeServiceNowtext(data: any): void {

    this.setState({ ServiceNowtext: data.target.value });

  }

  private changeReqName(data: any): void {

    this.setState({ ReqName: data.target.value });

  }

  private changeServiceNowTicketNumber(data: any): void {

    this.setState({ ServiceNowTicketNumber: data.target.value});

  }

  
  public handleRequestDateChange = (date: any) => {

    this.setState({ dtreqdate: date });

    }

    public handleReqDeliveryChange = (date: any) => {

      this.setState({ dtReqDelivery: date });

      }

      public handleServiceNowDateChange = (date: any) => {

        this.setState({ dtserviceNowdate: date });
  
        }


      public handledtServTickOpenChange = (date: any) => {

        this.setState({ dtServTickOpen: date });
            
        }

        private async getUserDetails()
        {

          

          let result= await this._service.getCurrentUser();
      
          this.setState({UserLoginName:result.Title});
              
        }

        //region PeoplePickeer and handleChange Events
  private async _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);

    if(items.length>0)
    {

      ClientPartnerName = items[0].text;

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info)=>{
      this.setState({userval:info});
      console.log(info);
 });

    }

    else
    {

      this.setState({userval:null});
    }

    //this.ppl.onChange([]);

  }

     private MyValidate()
     {

      

     }

     private Validations()
     {

      
      if(this.state.ReqName==null || this.state.ReqName=="")
      {
 
      alert('Please enter Requestor Name');
      this.setState({ flag: false });
      FinalStatus='';
      
 
     }
     else if(this.state.dtreqdate==null)
     {

      alert('Please select Requested Date');
      this.setState({ flag: false });
      FinalStatus='';
     }

     else if(this.state.dtReqDelivery==null || this.state.dtReqDelivery=="")
     {

      alert('Please select Request Delivery Date');
      this.setState({ flag: false });
      FinalStatus='';
     }

     
     else  if (this.state.System == null|| this.state.System=="") {

        alert('Please select System value');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if (this.state.Title == null || this.state.Title=="") {

        alert('Please enter Title of Change');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.Title.length>40)
      {
        
        alert('The title of change must be limited to 40 characters or less.');
        this.setState({ flag: false });
        FinalStatus='';

      }



      else if (this.state.Desc == null || this.state.Desc == "") {

        alert('Please enter Description');
        this.setState({ flag: false });
        FinalStatus='';
       
      }
   

      else if (this.state.ServiceNow == null || this.state.ServiceNow=="") 
      {

        alert('Please select if this change started as a ServiceNow Request');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

        if(this.state.ServiceNow=='Yes')
        {

          if(this.state.ServiceNowTicketNumber==null || this.state.ServiceNowTicketNumber=="")
          {
          alert('Please enter ServiceNow Ticket Number');
          this.setState({ flag: false });
          FinalStatus='';

          }

         

          else if(this.state.ServiceNowtext==null || this.state.ServiceNowtext=="")
          {

            alert('Please enter the Person Who Submitted the ServiceNow Ticket');
            this.setState({ flag: false });
            FinalStatus='';

          }

          else if(this.state.dtserviceNowdate==null)
          {

            alert('Please select date ServiceNow ticket was opened');
            this.setState({ flag: false });
            FinalStatus='';
          }

          //kalyanss
          
          if(FinalStatus=='Yes')
          {

            FinalStatus='Yes';
          }

          if(FinalStatus=="")
          {

            FinalStatus="";
          }

        }

        if(this.state.ServiceNow=='No')
        {

          this.setState({ flag: true });

          FinalStatus='Yes';

        }

        

     }

     private Clear()
     {

      this.setState({ ReqName: '' });

     }


      private OnBtnNextClick() :void {

      }

      private OnBtnSubmitClick() :void {

      if(this.state.ReqName==null || this.state.ReqName=="")
      {
 
      alert('Please enter Requestor Name');
      this.setState({ flag: false });
      FinalStatus='';
      
 
     }
     else if(this.state.dtreqdate==null)
     {

      alert('Please select Requested Date');
      this.setState({ flag: false });
      FinalStatus='';
     }

         
     else  if (this.state.System == null|| this.state.System=="") {

        alert('Please select System value');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if (this.state.Title == null || this.state.Title=="") {

        alert('Please enter Title of Change');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.Title.length>40)
      {
        
        alert('The title of change must be limited to 40 characters or less.');
        this.setState({ flag: false });
        FinalStatus='';

      }

      else if (this.state.Desc == null || this.state.Desc == "") {

        alert('Please enter Description');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.dtReqDelivery==null || this.state.dtReqDelivery=="")
      {
 
       alert('Please select Request Delivery Date');
       this.setState({ flag: false });
       FinalStatus='';
      }
 
   
      else if (this.state.ServiceNow == null || this.state.ServiceNow=="") 
      {

        alert('Please select if this change started as a ServiceNow Request');
        this.setState({ flag: false });
        FinalStatus='';
       
      }

        if(this.state.ServiceNow=='Yes')
        {

      if(this.state.ReqName==null || this.state.ReqName=="")
      {
 
     
      this.setState({ flag: false });
      FinalStatus='';
 
     }
     else if(this.state.dtreqdate==null)
     {

      
      this.setState({ flag: false });
      FinalStatus='';
     }

     
     
     else  if (this.state.System == null|| this.state.System=="") {

        
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if (this.state.Title == null || this.state.Title=="") {

       
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.Title.length>40)
      {
        
        
        this.setState({ flag: false });
        FinalStatus='';

      }

      else if (this.state.Desc == null || this.state.Desc == "") {

       
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.dtReqDelivery==null || this.state.dtReqDelivery=="")
     {

      
      this.setState({ flag: false });
      FinalStatus='';
     }

   
      else if (this.state.ServiceNow == null || this.state.ServiceNow=="") 
      {

        
        this.setState({ flag: false });
        FinalStatus='';
       
      }

      else if(this.state.ServiceNowTicketNumber==null || this.state.ServiceNowTicketNumber=="")
          {
          alert('Please enter ServiceNow Ticket Number');
          this.setState({ flag: false });
          FinalStatus='';

          }

         

          else if(this.state.ServiceNowtext==null || this.state.ServiceNowtext=="")
          {

            alert('Please enter the Person Who Submitted the ServiceNow Ticket');
            this.setState({ flag: false });
            FinalStatus='';

          }

          else if(this.state.dtserviceNowdate==null)
          {

            alert('Please select date ServiceNow ticket was opened');
            this.setState({ flag: false });
            FinalStatus='';
          }

          else
          {

            this.setState({ flag: true });

            FinalStatus='Yes';

            this.setState({ savedisabled: false });

            Disablevalue='Yes';

           

          }

        }

        if(this.state.ServiceNow=='No')
        {

          if(this.state.ReqName==null || this.state.ReqName=="")
          {
     
         
          this.setState({ flag: false });
          FinalStatus='';
     
         }
         else if(this.state.dtreqdate==null)
         {
    
          
          this.setState({ flag: false });
          FinalStatus='';
         }
    
         
    
         
         else  if (this.state.System == null|| this.state.System=="") {
    
            
            this.setState({ flag: false });
            FinalStatus='';
           
          }
    
          else if (this.state.Title == null || this.state.Title=="") {
    
           
            this.setState({ flag: false });
            FinalStatus='';
           
          }
    
          else if(this.state.Title.length>40)
          {
            
            
            this.setState({ flag: false });
            FinalStatus='';
    
          }
    
          else if (this.state.Desc == null || this.state.Desc == "") {
    
           
            this.setState({ flag: false });
            FinalStatus='';
           
          }

          else if(this.state.dtReqDelivery==null || this.state.dtReqDelivery=="")
         {
    
          
          this.setState({ flag: false });
          FinalStatus='';
         }

          else
          {

            this.setState({ flag: false });
            FinalStatus='Yes';
            this.setState({ savedisabled: false });
            Disablevalue='Yes';
           
          }
       
        

        }

        
        if(FinalStatus=="Yes")
        {

        FinalStatus='';

        if(this.state.ServiceNow=='Yes')
        {

          //#region 

let date=this.state.dtreqdate.getDate();

let month= (this.state.dtreqdate.getMonth()+1);

let year =(this.state.dtreqdate.getFullYear());

let FinalRequestDate=month+'/'+this.state.dtreqdate.getDate()+'/' +year;


let date1=this.state.dtReqDelivery.getDate();

let month1= (this.state.dtReqDelivery.getMonth()+1);

let year1 =(this.state.dtReqDelivery.getFullYear());

let FinalRequestDelDate=month1+'/'+this.state.dtReqDelivery.getDate() +'/' +year1;
        
let date2=this.state.dtserviceNowdate.getDate()

let month2= (this.state.dtserviceNowdate.getMonth()+1);

let year2 =(this.state.dtserviceNowdate.getFullYear());


let FinalServiceNowDate=month2 +'/'+this.state.dtserviceNowdate.getDate()+'/' +year2;


          let myfiles=[];

          for(var count=0;count<this.state.FileValue.length;count++)
          {
            
            myfiles.push(this.state.FileValue[count]);
          }

    let FilterTitle=this.state.Title.replace(/\s+$/, '');

    ReqSiteUrl=this.state.MyUrl;

    this._service.onDrop(this.state.MyColNameDesc,this.state.MyListName,this.state.ReqName,this.state.System,FilterTitle,this.state.Desc,this.state.ServiceNow,this.state.ServiceNowTicketNumber,FinalRequestDate,FinalRequestDelDate,FinalServiceNowDate,this.state.UserLoginName,this.state.ServiceNowtext,myfiles).then(function (data)
    {
      console.log(data);

      Disablevalue='Yes';

      alert('Record submitted successfully');
     
     window.location.replace(window.location.href);
      
    this.Clear();

      
    });

    }

    else if(this.state.ServiceNow=='No')
    {

     //#region 

let date=this.state.dtreqdate.getDate();

let month= (this.state.dtreqdate.getMonth()+1);

let year =(this.state.dtreqdate.getFullYear());


let FinalRequestDate=month+'/'+this.state.dtreqdate.getDate()+'/' +year;


let date1=this.state.dtReqDelivery.getDate()

let month1= (this.state.dtReqDelivery.getMonth()+1);

let year1 =(this.state.dtReqDelivery.getFullYear());


let FinalRequestDelDate=month1+'/'+this.state.dtReqDelivery.getDate() +'/' +year1;

//#endregion

   let myfiles=[];

    for(var count=0;count<this.state.FileValue.length;count++)
    {
            
    myfiles.push(this.state.FileValue[count]);

    }

    let FilterTitle1=this.state.Title.replace(/\s+$/, '');

    ReqSiteUrl=this.state.MyUrl;

    this._service.onDrop1(this.state.MyColNameDesc,this.state.MyListName,this.state.ReqName,this.state.System,FilterTitle1,this.state.Desc,this.state.ServiceNow,FinalRequestDate,FinalRequestDelDate,this.state.UserLoginName,myfiles).then(function (data)
    {
 
      Disablevalue='Yes';
      
      console.log(data);

      alert('Record submitted successfully');

      window.location.replace(window.location.href);

      
    });


        }
      }

      }


      private OnBtnBackClick() :void {

      }

    

    public ChangeSystem(ev: React.FormEvent<HTMLInputElement>, option: any): void {  

      this.setState({  

        System: option.key  
  
        });  

      }

      private changeFileupload(data: any) {

        let LocalFileVal= this.state.FileValue;
        
         LocalFileVal.push(data.target.files[0]);
        
        
        this.setState({FileValue:LocalFileVal});
        
        if(this.state.FileValue.length>4)
        {
        this.setState({disableFileUpload:true});
        
        }
        
        
        }

        private _removeItemFromDetail(Item: any) {
          console.log("itemId: " + Item.name); 
        
         let localFileValues=[];
        
         localFileValues=this.state.FileValue;
        
         if(localFileValues.length==1)
         {
        
          localFileValues=[];
         }
        
        
          for(var count=0;count<localFileValues.length;count++)
          {
        
            if(localFileValues[count].name==Item.name)
              {
                let Index=count;
        
                localFileValues.splice(Index,count);
        
              }
        
          }
        
          this.setState({FileValue:localFileValues,disableFileUpload:false});
        
        
        }

        public changeServiceNow=async(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void>=> {  
         
           this.setState({ServiceNow:option.key});
          
            if(option.key=='Yes')
            {

              
              this.setState({ServiceFlag:true });
              
            }

            if(option.key=='No')
            {

              
              this.setState({ServiceFlag:false });
            }


  
          }


  
  public render(): React.ReactElement<IIaAutomationVersion2Props> {

    //let filteredItems = this.filterListItems();

    return (
      <Stack tokens={stackTokens} styles={stackStyles} >
      <Stack>
        
        <p>Hi, <label>{this.state.UserLoginName}</label> When you submit this form, the owner will see your name and email address.</p><br></br>
        <p>Fields marked with an <label className={styles.redcolr}>*</label> are required. </p><br></br>
        <b><label className={styles.labelsFonts}>1. Requestor <label className={styles.redcolr}>*</label></label></b><br/>
        <p>Please enter an individuals first and last name or group name (Finance, Operations, etc.)</p>
        <div> 
        <input type="text" name="txtReqName" value={this.state.ReqName} onChange={this.changeReqName.bind(this)} className={styles.boxsize}/>
        
        </div><br/>
        <b><label className={styles.labelsFonts}>2. Request Date <label className={styles.redcolr}>*</label></label></b><br/>
        <div className={styles.datesize}> 
        <DateTimePicker  
          dateConvention={DateConvention.Date}  
          showLabels={false}
          value={this.state.dtreqdate}  
          onChange={this.handleRequestDateChange}
          
        />  

        </div> <br></br><br></br>
         <div> 
        <b><label className={styles.labelsFonts}>3. System <label className={styles.redcolr}>*</label></label></b><br></br><br></br>
        <ChoiceGroup className={styles.onlyFont}  id="System"  name="System" options={RadioSystem}   onChange={this.ChangeSystem.bind(this)}  selectedKey={this.state.System}/>
        </div> <br/>
        <b><label className={styles.labelsFonts}>4. Title of Change <label className={styles.redcolr}>*</label></label></b><br/>
        <div> 
       <input type="text" name="txtTitle" value={this.state.Title} onChange={this.changeTitle.bind(this)} className={styles.boxsize}/>
        </div><br/>
        <div>
          <ul>
            <li>
            Characters are limited to 40 characters or less.
            </li>
            <li>
            The following special characters are disabled: ? * \ / : & # " %
            </li>
          </ul>
       
       </div><br></br>

         <b><label className={styles.labelsFonts}> 5. Description <label className={styles.redcolr}>*</label></label></b><br/>
           <div>  
           <textarea id="txtDesc" value={this.state.Desc} onChange={this.changeDesc.bind(this)} className={styles.textAreacss}></textarea>
           </div><br/>

        <b><label className={styles.labelsFonts}>6. Request Delivery <label className={styles.redcolr}>*</label></label></b><br/>
        
        <p>Delivery dates can vary depending on the request and queue.</p><br></br>
        <div className={styles.datesize}> 
        <DateTimePicker  
          dateConvention={DateConvention.Date}  
          showLabels={false}
          value={this.state.dtReqDelivery}  
          onChange={this.handleReqDeliveryChange}
          
          />  

        
        </div> <br></br><br></br>

        <b><label className={styles.labelsFonts}>7. Upload any documents related to this request(Non-anonymous question)</label></b><br/>
        <input id="infringementFiles" type="file"  name="files[]"  onChange={this.changeFileupload.bind(this)} disabled={this.state.disableFileUpload}/>
        <br></br>
        <p>File number limit: 5Single file size limit: 10MBAllowed file types: Word, Excel, PPT, PDF, Image, Video, Audio</p>
          {this.state.FileValue.map((item,index) =>(

<div className={styles.padcss}>  
{item.name} <Icon iconName='Delete'  onClick={(event) => {this._removeItemFromDetail(item)}}/>
</div>
 
))}
 <Stack>  <br></br>
<b><label className={styles.labelsFonts}>8. Did this start as a ServiceNow Ticket ? <label className={styles.redcolr}>*</label></label></b><br></br>
<ChoiceGroup className={styles.onlyFont} options={RadioServiceNow}   onChange={this.changeServiceNow} />
</Stack><br></br>

{this.state.ServiceFlag == true &&
<Stack tokens={sectionStackTokens}>
<div id="divNext">
{/* <PrimaryButton text="Next" onClick={this.OnBtnNextClick.bind(this)} styles={stackButtonStyles} className={styles.Mybutton}/> */}
</div><br></br>
<b><label className={styles.labelsFonts}>ServiceNow Ticket Information</label></b>
<br></br>
<label className={styles.labelsFonts}>Complete this section if the request started as a ServiceNow ticket to the help desk.</label><br></br>
<b><label className={styles.labelsFonts}>9. ServiceNow Ticket Number ?  <label className={styles.redcolr}>*</label></label></b><br/>
<input type="text" name="txtServiceNumber" value={this.state.ServiceNowTicketNumber} onChange={this.changeServiceNowTicketNumber.bind(this)} className={styles.boxsize}/><br></br>
<b><label className={styles.labelsFonts}>10. Date ServiceNow Ticket Opened  <label className={styles.redcolr}>*</label></label></b><br/>
<div className={styles.datesize}> 
{/* <DatePickerComponent  value={this.state.dtServTickOpen} onChange={this.handledtServTickOpenChange} ></DatePickerComponent><br/> */}

          <DateTimePicker  
          dateConvention={DateConvention.Date}  
          showLabels={false}
          value={this.state.dtserviceNowdate}  
          onChange={this.handleServiceNowDateChange}
          
          />  
</div> <br/>
<b><label className={styles.labelsFonts}>11. Person Who Submitted the Ticket ? <label className={styles.redcolr}>*</label></label></b><br></br>
<div>  
{/* <PeoplePicker 
                context={this.props.context}
                //titleText="User Name"
                personSelectionLimit={1}
                showtooltip={true}
                required={true}
                disabled={false}
                onChange={this._getPeoplePickerItems.bind(this)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                defaultSelectedUsers={(this.state.ClientPatnerName && this.state.ClientPatnerName.length) ? [this.state.ClientPatnerName] : []}
                ref={c => (this.ppl = c)} 
                resolveDelay={1000} />   */}

<input type="text" name="txtServiceNowtext" value={this.state.ServiceNowtext} onChange={this.changeServiceNowtext.bind(this)} className={styles.boxsize}/>
</div><br></br>

{/* <input type="text" name="txtEmail" value={this.state.Email} onChange={this.changeEmail.bind(this)} className={styles.boxsize}/><br></br> */}
<Stack horizontal tokens={sectionStackTokens}>
<StackItem className={styles.commonstyle}> 
{/* <PrimaryButton text="Back" onClick={this.OnBtnBackClick.bind(this)} styles={stackButtonStyles} className={styles.Mybutton}/> */}
</StackItem>

</Stack>
</Stack>
}

<StackItem className={styles.commonstyle}> 
<PrimaryButton text="Submit" onClick={this.OnBtnSubmitClick.bind(this)} styles={stackButtonStyles} className={styles.Mybutton} disabled={Disablevalue==null?false :true }/>
</StackItem>

</Stack>
</Stack>
         
    );
  }
}
