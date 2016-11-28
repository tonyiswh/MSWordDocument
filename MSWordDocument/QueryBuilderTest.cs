using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ActiveQueryBuilder.Core;
using ActiveQueryBuilder.View.WinForms;
using System.Data.OleDb;

namespace MSWordDocument
{
    public partial class QueryBuilderTest : Form
    {
       
        public QueryBuilderTest()
        {
            InitializeComponent();

            queryBuilder1 = new QueryBuilder();

            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = @"Provider=SQLNCLI11;Server=DWS02;Database=AdventureWorks2012;Integrated Security=SSPI;";

            OLEDBMetadataProvider metaProvider = new OLEDBMetadataProvider();
            metaProvider.Connection = connection;

            GenericSyntaxProvider syntaxProvider = new GenericSyntaxProvider();

            queryBuilder1.MetadataProvider = metaProvider;
            queryBuilder1.SyntaxProvider = syntaxProvider;

            queryBuilder1.InitializeDatabaseSchemaTree();
        }

        private void QueryBuilder_Load(object sender, EventArgs e)
        {

        }


    }
}
