using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using com.netsuite.webservices;

//for Local DB 
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;


/* This is our new lead class - the class will have methods for accepting data and will apply business rules as needed */
public partial class _Default : System.Web.UI.Page {
    class New_Lead {
        private NetSuiteService _service;
        private Boolean _isAuthenticated;

        private String _companyName;
        private String _email;
        private String _phone;
        private String _firstname;
        private String _lastname;
        private String _industry;
        private String _source;
        private String _ip;

        private String _customerID;
        private String _entityName;
        private String _contactID;

        private Status _customerStatus;
        private Status _contactStatus;
        private Status _attachStatus;

        public New_Lead() {
			_isAuthenticated = false;
			_service = new NetSuiteService();
			_service.Timeout = 1000*60*60*2;

            //Enable cookie management
            Uri myUri = new Uri("http://www.domain.com");
            _service.CookieContainer = new CookieContainer();
        }

        public void login(string email, string password, string account) {
            //login to NetSuite
            Passport passport = new Passport();
            passport.account = account;
            passport.email = email;
            passport.password = password;
            Status status = _service.login(passport).status;
            _isAuthenticated = status.isSuccess;
        }

        public void logout() {
            _isAuthenticated = !(_service.logout().status.isSuccess);
        }

        public Boolean isAuthenticated() {
            return _isAuthenticated;
        }

        public void sendEmail() {
            MailAddress from = new MailAddress("sales@domain.com", "Sales Department");
            MailAddress to = new MailAddress(_email, _firstname + " " + _lastname);
            MailMessage message = new MailMessage(from,to);
            String body; 
            using (StreamReader sr = new StreamReader("path_to_email_template.html")) {
                body = sr.ReadToEnd();
            }
            //body = "This is a test.<br>";
            message.IsBodyHtml = true;
            message.Subject = "Thank You";
            message.Body = body;
            SmtpClient client = new SmtpClient("smtp.domain.com");
            client.Send(message);
        }

        /*
            At one point, NetSuite's WebServices was considered to be buggy.  As a backup, the client asked that this form save
            each submission to a local database.  This is no longer needed but is left in as an example of ODBC connectivity
        */
        public void saveToDB() {
            string sDBConn = "Data Source=DB_SERVER,1433;Initial Catalog=leads;User ID=sa;Password=1234";
            string query = "INSERT INTO xx_NeedsAnalysis VALUES(@lead_datetime,@lead_lastname,@lead_firstname,@lead_companyname,@lead_email,@lead_phone,@lead_industry,@lead_remoteip,@lead_source)";

            SqlConnection conn = new SqlConnection(sDBConn);
            SqlCommand cmd = new SqlCommand(query, conn);

            //BIND THE DATA TO THE PARAMETERS
            cmd.Parameters.AddWithValue("@lead_datetime", DateTime.Now);
            cmd.Parameters.AddWithValue("@lead_lastname", _lastname);
            cmd.Parameters.AddWithValue("@lead_firstname", _firstname);
            cmd.Parameters.AddWithValue("@lead_companyname",_companyName);
            cmd.Parameters.AddWithValue("@lead_email",_email);
            cmd.Parameters.AddWithValue("@lead_phone",_phone);
            cmd.Parameters.AddWithValue("@lead_industry", _industry);
            cmd.Parameters.AddWithValue("@lead_remoteip", _ip);
            cmd.Parameters.AddWithValue("@lead_source", _source);

            //OPEN THE CONNECTION AND EXECUTE THE QUERY
            cmd.Connection.Open();
            cmd.ExecuteNonQuery();
        }

        /*
            These three calls parse the returned XML from NetSuite to determine if the operations for creating
            a new customer, creating a new contact, and attaching the new contact to the new customer were a
            success, respectively.
        */
        public String getCustomerStatus() {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (_customerStatus.statusDetail != null) {
                for (int i = 0; i < _customerStatus.statusDetail.Length; i++) {
                    sb.Append("[Code=" + _customerStatus.statusDetail[i].code + "] " + _customerStatus.statusDetail[i].message + "\n");
                }
            }
            return sb.ToString();
        }

        public String getContactStatus() {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (_contactStatus.statusDetail != null) {
                for (int i = 0; i < _contactStatus.statusDetail.Length; i++) {
                    sb.Append("[Code=" + _contactStatus.statusDetail[i].code + "] " + _contactStatus.statusDetail[i].message + "\n");
                }
            }
            return sb.ToString();
        }

        public String getAttachStatus() {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (_attachStatus.statusDetail != null) {
                for (int i = 0; i < _attachStatus.statusDetail.Length; i++) {
                    sb.Append("[Code=" + _attachStatus.statusDetail[i].code + "] " + _attachStatus.statusDetail[i].message + "\n");
                }
            }
            return sb.ToString();
        }

        public void loadData(String companyName, String email, String phone, String firstname, String lastname, String industry, String ip, String source) {
            if (companyName == null) { companyName = ""; }
            _companyName = companyName;

            if (firstname == null) { firstname = ""; }
            _firstname = firstname;

            if (lastname == null) { lastname = ""; }
            _lastname = lastname;

            if (email == null) { email = ""; }
            _email = email;

            if (phone == null) { phone = ""; }
            _phone = phone;

            if (industry == null) { industry = ""; }
            _industry = industry;

            if (source == null) { source = ""; }
            _source = source;

            if (ip == null) { ip = ""; }
            _ip = ip;
        }

        /* 
            Creates the new Lead/Customer record in NetSuite.  Must be called first because we'll need the resulting Customer ID
            to place in the new Contact record, and both ids to run the 'attach' operation.  These are NetSuite's rules, sadly. 
        */

        public void createCustomer() {
            Customer customer = new Customer();

            // Populate information provided by the user
            customer.companyName = _companyName;
            customer.email = _email;
            customer.phone = _phone;

            // Set as Lead - New
            RecordRef status = new RecordRef();
            status.internalId = "6"; // LEAD - New
            customer.entityStatus = status;

            /*
                Custom Fields require creating a reference record based on the type of field, and then we
                add each reference object to an array that gets passed into the XML renderer that NetSuite
                provides for us.

                Industry, IP address and Source are custom fields we've created.  Industry is a drop down list, which
                means our reference object must contain a reference to which drop down entry the new lead selected.
                This seems a bit overboard but lists can be multi-selects as well.  IP Address and source are just simple
                text fields, so nothing fancy is required, they're just a simple name/value pair.
            */

            System.Collections.ArrayList customFields = new System.Collections.ArrayList();

            if (_industry != "") {
                SelectCustomFieldRef custIndustry = new SelectCustomFieldRef();
                custIndustry.internalId = "custentity_customerindustry";
                ListOrRecordRef custIndustryList = new ListOrRecordRef();
                custIndustryList.internalId = _industry;
                custIndustry.value = custIndustryList;
                customFields.Add(custIndustry);
            }

            if (_ip != "") {
                StringCustomFieldRef ip = new StringCustomFieldRef();
                ip.internalId = "custentity_fnaipaddress";
                ip.value = _ip;
                customFields.Add(ip);
            }

            if (_source != "") {
                StringCustomFieldRef source = new StringCustomFieldRef();
                source.internalId = "custentity_custfnasource";
                source.value = _source;
                customFields.Add(source);
            }

            CustomFieldRef[] customFieldRefs = new CustomFieldRef[customFields.Count];
            IEnumerator ienum = customFields.GetEnumerator();
            for (int i = 0; ienum.MoveNext(); i++) {
                customFieldRefs[i] = (CustomFieldRef)ienum.Current;
            }
            customer.customFieldList = customFieldRefs;

            /* 
                Now with our customer record created, we're going to add it to our XML renderer and pass it on NetSuite.
                NetSuite will be returning a response in XML so we're going to grab that so we can get the id of the newly
                created record in NetSuite.
            */
            WriteResponse response = _service.add(customer);
            _customerStatus = response.status;
            if (this.getCustomerStatus() == "") {
                _customerID = ((RecordRef)response.baseRef).internalId;
                _entityName = ((RecordRef)response.baseRef).name;
            }
        }

        /*
            If createCustomer() has been called, we can create the contact.  We could actually create a new contact without
            a customer id - but the problem is that with so many contacts, this would be hard to find.  Plus, contacts don't
            really have a CRM status in NetSuite like customers do.  We want our sales teams to be alerted for LEAD - New.
        */
        public void createContact() {
            Contact contact = new Contact();

            // Populate information provided by the user
            contact.firstName = _firstname;
            contact.lastName = _lastname;
            contact.email = _email;
            contact.unsubscribe = false;
            contact.unsubscribeSpecified = true;
            contact.phone = _phone;
            contact.company = new RecordRef();
            contact.company.internalId = _customerID;
            contact.company.name = _entityName;

            // Add the new contact to the database
            WriteResponse response = _service.add(contact);
            _contactStatus = response.status;
            if (this.getContactStatus() == "") {
                _contactID = ((RecordRef)response.baseRef).internalId;
            }
        }

        /*
            Attaching the contact to the customer is not only establishing the 1-to-N relationship but also
            assigning a role to the contact within the customer organization.  This form doesn't ask what the
            end user's role is within their organization, but it could and that data could be recorded here.
            Right now we're assuming that the customer is the primary point of contact.
        */
        public void attachContact() {
            // Create a new contact role (primary)
            RecordRef contactRole = new RecordRef();
            contactRole.internalId = "-10";  //Primary
            contactRole.type = RecordType.contactRole;
            contactRole.typeSpecified = true;

            /* We're going to create a reference to the new contact record via it's internal ID */
            RecordRef contactRef = new RecordRef();
            contactRef.internalId = _contactID;
            contactRef.type = RecordType.contact;
            contactRef.typeSpecified = true;

            /* And the same here for the newly created customer record */
            RecordRef customerRef = new RecordRef();
            customerRef.internalId = _customerID;
            customerRef.type = RecordType.customer;
            customerRef.typeSpecified = true;

            /* We attach the two by using a pre-made XML renderer from NetSuite. */
            AttachContactReference attachReference = new AttachContactReference();
            attachReference.attachTo = customerRef;
            attachReference.contact = contactRef;
            attachReference.contactRole = contactRole;

            _attachStatus = _service.attach(attachReference).status;
        }

    }

    //debug
    public String getStatus(Status status) {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        if (status.statusDetail != null) {
            for (int i = 0; i < status.statusDetail.Length; i++) {
                sb.Append("[Code=" + status.statusDetail[i].code + "] " + status.statusDetail[i].message + "\n");
            }
        }
        return sb.ToString();
    }

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void Save_Lead() {
        New_Lead submitted_lead = new New_Lead();
        String companyName;
        String email;
        String phone;
        String firstname;
        String lastname;
        String industry;
        String source;
        String ip;

        companyName = Request.Form.Get("txtOrganization");
        email = Request.Form.Get("txtEmail");
        phone = Request.Form.Get("txtPhone");
        firstname = Request.Form.Get("txtFirstname");
        lastname = Request.Form.Get("txtLastname");
        industry = Request.Form.Get("txtNeedsIndustry");
        source = Request.Form.Get("source");
        ip = Request.ServerVariables.Get("REMOTE_ADDR");

        submitted_lead.loadData(companyName, email, unsubscribe, phone, firstname, lastname, industry, ip, source);
        submitted_lead.saveToDB();
        submitted_lead.login("netsuiteuser@domain.com", "password", "012345"); // user, pass & netsuite id
        submitted_lead.createCustomer();
        submitted_lead.createContact();
        submitted_lead.attachContact();
        submitted_lead.logout();

        if (email != "" && email != null) {
            submitted_lead.sendEmail();
        }
    }

    
}
