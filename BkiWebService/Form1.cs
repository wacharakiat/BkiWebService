using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Configuration;
using BkiWebService.com.bangkokinsurance.www;

namespace BkiWebService
{
    public partial class frmSend : Form
    {
        public frmSend()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string thisdate = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString() + "0" + DateTime.Today.Day.ToString();
            txtSeqNo.Text = "HBC" + thisdate;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            string userId = "Claimoutsrc";
            string password = "ClmOutSrcP@ssw0rd";
            string connectionString = ConfigurationManager.ConnectionStrings["ClaimOle"].ToString();

            //สร้าง Master Table
            string selectCommandText = "select * from bki_clmhead_reimb";
            OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommandText, connectionString);
            OleDbCommandBuilder cb = new OleDbCommandBuilder(adapter);
            DataTable masTable = new DataTable { TableName = "clm_mas" };
            adapter.Fill(masTable);

            masTable.Columns["claim_st"].ColumnName = "claim_status";
            masTable.Columns["paye_addr1"].ColumnName = "payee_addr1";
            masTable.Columns["paye_addr2"].ColumnName = "payee_addr2";
            masTable.Columns["payee_jw"].ColumnName = "payee_jw_code";
            masTable.Columns["payee_zip"].ColumnName = "payee_zip_code";
            masTable.Columns["bank_br"].ColumnName = "bank_br_code";
            masTable.Columns["bank_accno"].ColumnName = "bank_acc_no";
            masTable.Columns["bki_clm_no"].ColumnName = "batch_no";
            masTable.Columns["clmdecline"].ColumnName = "clm_decline";
            masTable.AcceptChanges();

            // สร้าง Detail Table
            string selectCommandText1 = "select * from bki_clmdetail_reimb";
            OleDbDataAdapter detAdapter = new OleDbDataAdapter(selectCommandText1, connectionString);
            OleDbCommandBuilder detCb = new OleDbCommandBuilder(detAdapter);
            DataTable detTable = new DataTable { TableName = "clm_detail" };
            detAdapter.Fill(detTable);
            if (masTable.Rows.Count > 0 && detTable.Rows.Count > 0)
            {
                DataTable result = new DataTable { TableName = "clm_result" };
                send_clm_ostdata_prod send = new send_clm_ostdata_prod();
                result = send.send_clm_ostdata(txtSeqNo.Text, userId, password, masTable, detTable);
                txtResult.Text = result.Rows[0]["return_descr"].ToString();
            }
            else
            {
                txtResult.Text = "no data for send";
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            this.txtFileName.Text = "File is here";
        }

    }
}