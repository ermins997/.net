using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraTreeList.Columns;
using IvisTP300;
using IvisXB300.Models;
using IvisXP300;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static IvisTP300.tT1P;

namespace T8IN300
{
    public partial class T8INF : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        FileSystemWatcher watcher;
        List<extendedt1> mGlobalNormativ;  
        List<xSiIdent> identi;
        List<xSiIdent> identiZaDodati;
        List<extendedt1> MaterijaliBezSifara;
        bool mblnbarEditItemOverwriteT;
        bool mblnAutoSave;
        bool mblnNotifyNOK;
        bool mblnNotifyOK;

        delegate void SetTextCallback(List<extendedt1> lNormativ);

        public T8INF()
        {
            InitializeComponent();
            try
            {

                mGlobalNormativ = new List<extendedt1>(); 
                identiZaDodati = new List<xSiIdent>();
                MaterijaliBezSifara = new List<extendedt1>();
        
                repositoryItemSearchLookUpEditID.Properties.PopupFormSize = new Size(1000, 300);
                barEditItemFolder.EditValue = Properties.Settings.Default["Location"];
                barEditItemFormat.EditValue = Properties.Settings.Default["Mode"];
                barEditItemAutoSave.EditValue = Properties.Settings.Default["AutoSave"];
                barEditItemOverwriteT.EditValue = Properties.Settings.Default["OverwriteT"];
                barEditItemNotifyNOK.EditValue = Properties.Settings.Default["NotifyNOK"];
                barEditItemNotifyOK.EditValue = Properties.Settings.Default["NotifyOK"];

                ReadFolder();

                barEditItemRefresh.EditValue = Properties.Settings.Default["RefreshTime"];
                StartWatching();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void barEditItemFolder_ItemPress(object sender, ItemClickEventArgs e)
        {
            setLocationDirectory();
        }

        private void setLocationDirectory()
        {
            string result = "";
            while (result == "")
            {
                xtraFolderBrowserDialog.SelectedPath = Properties.Settings.Default["Location"].ToString();
                xtraFolderBrowserDialog.Title = "Izaberite direktorij u kom se nalaze podaci za uvoz";
                xtraFolderBrowserDialog.ShowDialog();
                result = xtraFolderBrowserDialog.SelectedPath;
            }
            barEditItemFolder.EditValue = result;

            if (!Directory.Exists(result + "\\Processed"))
                Directory.CreateDirectory(result + "\\Processed");

            Properties.Settings.Default["Location"] = barEditItemFolder.EditValue;
            Properties.Settings.Default.Save();
        }

        private void barEditItemFormat_EditValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["Mode"] = barEditItemFormat.EditValue;
            Properties.Settings.Default.Save();
        }

        private void barButtonItemExit_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (MessageBox.Show("Želite li zatvoriti T8IM?", "Izlaz", MessageBoxButtons.YesNo) == DialogResult.Yes)
                Application.Exit();
        }

        private void ReadFolder()
        {
            try
            {
                splashScreenManagerW.ShowWaitForm();
                timerrefreshid.Enabled = false;

                DirectoryInfo d = new DirectoryInfo(barEditItemFolder.EditValue.ToString());
                FileInfo[] Files = d.GetFiles("*.xml");
                foreach (FileInfo file in Files)
                {
                    ReadXML(file.FullName);
                }

                SetDataSource(mGlobalNormativ);
                tlistNormativ.RefreshDataSource();

                if (mblnAutoSave)
                    SaveToIvis();

                timerrefreshid.Enabled = true;
                if (splashScreenManagerW.IsSplashFormVisible)
                    splashScreenManagerW.CloseWaitForm();
            }
            catch (DirectoryNotFoundException ex)
            {
                splashScreenManagerW.CloseWaitForm();
                setLocationDirectory();
            }
        }

        private void CleanFolder()
        {
            string sourceDirectory;
            string destinationDirectory;
            string destinationFile;
            string initialdestinationFile;
            int counter = 1;

            sourceDirectory = barEditItemFolder.EditValue.ToString();
            destinationDirectory = barEditItemFolder.EditValue + "\\Processed";
            if (!Directory.Exists(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);

            DirectoryInfo d = new DirectoryInfo(sourceDirectory);

            foreach (extendedt1 currt1 in mGlobalNormativ.Where(x => x.ExistAllMaterialIDS).ToList())
            {
                FileInfo file = d.GetFiles(Path.GetFileName(currt1.FullFilename)).First();
                destinationFile = destinationDirectory + "\\" + file.Name;
                initialdestinationFile = destinationFile;
                if (File.Exists(destinationFile))
                {
                    var idx = destinationFile.LastIndexOf(".");
                    if (idx > -1)
                        destinationFile = destinationFile.Substring(0, idx) + "_" + counter.ToString("000") + destinationFile.Substring(idx);

                    while (File.Exists(destinationFile))
                    {
                        counter += 1;
                        destinationFile = initialdestinationFile.Substring(0, idx) + "_" + counter.ToString("000") + initialdestinationFile.Substring(idx);
                    }
                }
                File.Move(sourceDirectory + "\\" + file.Name, destinationFile);

            }
        }

        private void StartWatching()
        {
            watcher = new FileSystemWatcher(barEditItemFolder.EditValue.ToString());

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Created += OnCreated;
            watcher.Error += OnError;

            watcher.Filter = "*.xml";
            watcher.IncludeSubdirectories = false;
            watcher.EnableRaisingEvents = true;
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            Task.WaitAll();

            splashScreenManagerW.ShowWaitForm();
            timerrefreshid.Enabled = false;
            string value = e.FullPath;
            ReadXML(value);

            SetDataSource(mGlobalNormativ);

            if (mblnAutoSave)
                SaveToIvis();
            timerrefreshid.Enabled = true;
            timerrefreshid.Enabled = true;

            splashScreenManagerW.CloseWaitForm();
        }

        private void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException());

        private void PrintException(Exception ex)
        {
            if (ex != null)
            {
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }

        private void ReadXML(string path)
        {

            bool timerrefreshidenabled = timerrefreshid.Enabled;
            timerrefreshid.Enabled = false;

            try
            {
                XmlDocument doc = new XmlDocument();
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
                doc.Load(path);
                Application.DoEvents();

                if (barEditItemFormat.EditValue.ToString().Equals("1 Elas PDM XML"))
                {
                  
                    foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                    {

                        XmlNode transactionnode = node.FirstChild;
                        XmlNode documentnode = transactionnode.FirstChild;

                        ReadNode(documentnode.FirstChild, null);

                        mGlobalNormativ.Last().ExistAllMaterialIDS = ProvjeriPostojanjeSifara(mGlobalNormativ.Last()); 
                        mGlobalNormativ.Last().FullFilename = path;
                    }
                } else 
                {
                    XmlNode nodeLogi = doc.DocumentElement;
                 
                    ReadParentDiscountNode(nodeLogi);


                    foreach (XmlNode node1 in doc.DocumentElement.ChildNodes)
                    {

                        if (node1.Name == "Position")
                        {
                            foreach (XmlNode node2 in node1.ChildNodes)
                            {
                                if (node2.Name == "Profiles")
                                {

                                    foreach (XmlNode node3 in node2.ChildNodes)
                                    {
                                        ReadDiscountNode(node3, mGlobalNormativ.Last()); 
                                    }
                                }

                            }
                        }
                        
                        if (node1.Name == "ArticleSummary")
                        {
                            foreach (XmlNode node2 in node1.ChildNodes)
                            {
                                if (node2.Name == "AllFittings" || node2.Name == "AllAccessories")
                                {
                                    foreach (XmlNode node3 in node2.ChildNodes)
                                    {
                                        ReadDiscountNode(node3, null); 
                                    }
                                }

                            }
                        }
                        

                    }
                }
                MaterijaliBezSifara = MaterijaliBezSifara.GroupBy(p => new { p.vrsta, p.debljina }).Select(x => x.First()).ToList();
                Application.DoEvents();
                timerrefreshid.Start();
            }
            catch (Exception exc)
            {
                Console.WriteLine("Nije udure: " + exc.Message);
                throw;

            }
            timerrefreshid.Enabled = timerrefreshidenabled;

        }

        private void SetDataSource(List<extendedt1> lNormativ)
        {

            if (this.tlistNormativ.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetDataSource);
                this.Invoke(d, new object[] { lNormativ });
            }
            else
            {
                LoadrepositoryItemSearchLookUpEditID();
                repositoryItemSearchLookUpEditID.DataSource = identi.Concat(identiZaDodati).ToList();
                tlistNormativ.DataSource = lNormativ;
                tlistNormativ.RefreshDataSource();
                tlistNormativ.ExpandAll();
                tlistNormativ.Enabled = true;
            }
        }

        private void ReadParentDiscountNode(XmlNode parentNode)
        {
            extendedt1 id = new extendedt1();
            id.ExistAllMaterialIDS = true;
            id.jm = "kom";
            id.kp = "MS";
            id.Saved = false;
            id.ks = parentNode.Attributes[1].Value.Substring(0, 1);
            id.no = parentNode.Attributes[1].Value;
           


          
            mGlobalNormativ.Add(id);

        }
        private void ReadDiscountNode(XmlNode node, extendedt1 parent)
        {

            
            {
                xSiIdents lxSiIdents = new xSiIdents(); 

                var number = node.Attributes["Number"].Value;
                var name = node.Attributes["Name"].Value;
                var qty = "";
                if (node.Attributes["Qty"] != null)
                {
                   
                if (node.Attributes["Qty"].Value != null)
                {
                     qty = node.Attributes["Qty"].Value;
                } 
                }

               
                if (name != null && number != null)
                {
                    xSiIdent newxSiIdent = new xSiIdent();
                    extendedt1 t1_id_objekat = new extendedt1();

                    t1_id_objekat.ExistAllMaterialIDS = true;
                    t1_id_objekat.jm = "kom";
                    t1_id_objekat.kp = "MS";
                    t1_id_objekat.Saved = false;

                    t1_id_objekat.t1id = number;
                    t1_id_objekat.nd = name;
                    t1_id_objekat.no = name;
                    t1_id_objekat.t1kl = qty == "" ? 0 : Convert.ToDouble(qty);

                    newxSiIdent.sinz = id.kp;
                    newxSiIdent.siqt = id.ks;
                    newxSiIdent.siid = id.t1id;
                    newxSiIdent.sino = id.no;
                    newxSiIdent.sind = id.nd;
                    newxSiIdent.sijm = id.jm;
                    newxSiIdent.sidt = DateTime.Now;

                    xSiIdent lxSiIdent = new xSiIdent();
                    try
                    {
                        
                        lxSiIdent = lxSiIdents.GetSiByAPandID("31115", t1_id_objekat.t1id);
                    }
                    catch { }

                    if (lxSiIdent == null || lxSiIdent.siid == null)
                    {
                        t1_id_objekat.exists = false;

                        xSiIdent newxSiIdent = new xSiIdent();
                        newxSiIdent.sikp = t1_id_objekat.kp;
                        newxSiIdent.siks = t1_id_objekat.ks; 
                        newxSiIdent.siid = t1_id_objekat.t1id; 
                        newxSiIdent.sino = t1_id_objekat.no; 
                        newxSiIdent.sind = t1_id_objekat.nd;
                        newxSiIdent.sijm = t1_id_objekat.jm; 
                        newxSiIdent.sidt = DateTime.Now;

                        identiZaDodati.Add(newxSiIdent); 
                    }
                    else
                    {
                        t1_id_objekat.exists = true;
                    }

                    


                    lxSiIdent = null;
                    
                    if (parent == null)
                        mGlobalNormativ.Add(t1_id_objekat); 
                        else
                        parent.Children.Add(t1_id_objekat);
                    Console.WriteLine("Name is " + " name is " + name + " number is" + number + " length is " + qty);
                    Console.ReadLine();
                   
                }
            }














            var name = node.Attributes["Name"].Value;
            var price = node.Attributes["Price"].Value;
            var number = node.Attributes["Number"].Value;
            var length = node.Attributes["Length"].Value;


            var description = node.FirstChild.Attributes["Description"].Value;ovo se nalazi u slijedecem nodu i odnosi se na boju jer se node zove Color
            var price = node.Attributes["Price"].Value;
            var number = node.Attributes["Number"].Value;
            var qty = node.Attributes["Qty"].Value;

            
            if (price != null && number != null && qty != null)
            {
                Console.WriteLine("Name is " + " price is " + price + " number is" + number + " length is " + qty);
                Console.ReadLine();
                 Process the value here
            }
            xSiIdents lxSiIdents = new xSiIdents();
            extendedt1 id = new extendedt1();
            id.Children = new List<extendedt1>();

            id.ExistAllMaterialIDS = true;
            id.jm = "kom";
            id.kp = "MS";
            id.Saved = false;



        } 

        private void ReadNode(XmlNode node, extendedt1 parent)
        {
            string lstrVM = string.Empty;
            string lstrDL = string.Empty;
            double ldblKL = 0;
            double ldblDL = 0;

            xSiIdents lxSiIdents = new xSiIdents();
            extendedt1 id = new extendedt1();
            id.Children = new List<extendedt1>();


            #region priprema t1 objekta
            id.ExistAllMaterialIDS = true;
            id.jm = "kom";
            id.kp = "MS";
            id.Saved = false;
            if (parent != null)
            {
                id.t1idm = parent.t1id;
                id.t1tsm = parent.t1ts;
            }

            foreach (XmlNode atributenode in node.ChildNodes)
            {
                if (atributenode.Name == "attribute")
                {

                    if (atributenode.Attributes.Count == 2)
                    {
                        switch (atributenode.Attributes[0].Value)
                        {
                            case "Naziv":
                                id.ks = atributenode.Attributes[1].Value.Substring(0, 1);
                                id.no = atributenode.Attributes[1].Value;
                                break;

                            case "Name":
                                id.nd = atributenode.Attributes[1].Value;
                                break;

                            default:
                                switch (id.ks)
                                {
                                    case "P":
                                    case "D":
                                        if (atributenode.Attributes[0].Value == "Ident IVIS - P")
                                            if (!string.IsNullOrEmpty(atributenode.Attributes[1].Value))
                                            {
                                                id.t1id = id.ks + atributenode.Attributes[1].Value;
                                                id.icon = 1;
                                            }

                                        break;

                                    case "S":
                                        if (atributenode.Attributes[0].Value == "Ident IVIS - S")
                                            if (!string.IsNullOrEmpty(atributenode.Attributes[1].Value))
                                            {
                                                id.t1id = id.ks + atributenode.Attributes[1].Value;
                                                id.icon = 0;
                                            }
                                        break;
                                }

                                break;

                            case "Vrsta Materijala":
                                lstrVM = atributenode.Attributes[1].Value;
                                break;

                            case "Debljina Lima":
                                lstrDL = atributenode.Attributes[1].Value;
                                break;
                        }
                    }
                }

                else if (atributenode.Name == "references")
                {
                    foreach (XmlNode documentnode in atributenode.ChildNodes)
                    {
                        ReadNode(documentnode.FirstChild, id);
                    }
                }
            }

            id.t1te = id.ks;
            id.t1tsm = "IZ";
            id.t1ts = "IZ";
            try
            {
                id.t1kl = double.Parse(node.Attributes[1].Value);
            }
            catch { }
            #endregion priprema t1 objekta


          
            xSiIdent lxSiIdent = new xSiIdent();
            try
            {
                
                lxSiIdent = lxSiIdents.GetSiByAPandID("31115", id.t1id);
            }
            catch { }

            if (lxSiIdent == null || lxSiIdent.siid == null)
            {
                id.exists = false;

                xSiIdent newxSiIdent = new xSiIdent();
                newxSiIdent.sikp = id.kp;
                newxSiIdent.siks = id.ks; 
                newxSiIdent.siid = id.t1id; 
                newxSiIdent.sino = id.no; 
                newxSiIdent.sind = id.nd; 
                newxSiIdent.sijm = id.jm; 
                newxSiIdent.sidt = DateTime.Now;

                identiZaDodati.Add(newxSiIdent);
            }

            else
            {
                id.exists = true;
            }


            lxSiIdent = null;

            if (parent == null)
                mGlobalNormativ.Add(id); 
            else
                parent.Children.Add(id);

           

            if (!string.IsNullOrEmpty(lstrVM) || !string.IsNullOrEmpty(lstrDL))
            {
                extendedt1 newT1 = new extendedt1();
                newT1.t1idm = id.t1id;
                newT1.t1tsm = "IZ";
                newT1.t1rb = 1;
                newT1.no = lstrVM + " " + lstrDL;



                newT1.nd = "Nema u šifarniku";
                newT1.icon = 2;
                newT1.vrsta = lstrVM;
                newT1.debljina = lstrDL;
                newT1.exists = false;

                newT1.debljina = lstrDL;
                if (IsNumeric(lstrDL))
                {
                    ldblDL = double.Parse(lstrDL) / 10;
                    lstrDL = ldblDL.ToString().Replace(",", ".");
                    newT1.debljina = lstrDL;

                }
                if (!string.IsNullOrEmpty(lstrVM) && (ldblDL > 0 || lstrDL == ""))
                {

                    List<xSiIdent> idList = lxSiIdents.GetSiBySivaSivf("31115", lstrVM, lstrDL);
                    if (idList.Count > 0)
                    {
                        newT1.t1id = idList.FirstOrDefault().siid;
                        newT1.ks = idList.FirstOrDefault().siks;
                        newT1.kp = idList.FirstOrDefault().sikp;
                        newT1.no = idList.FirstOrDefault().sino;
                        newT1.nd = idList.FirstOrDefault().sind;
                        newT1.jm = idList.FirstOrDefault().sijm;
                        newT1.exists = true;
                        newT1.ExistAllMaterialIDS = true;
                    }
                    else
                    {
                        newT1.ExistAllMaterialIDS = false;
                    }
                }

                newT1.t1ts = "MT";
                newT1.t1te = "R";
                newT1.t1kl = ldblKL;
                newT1.vrsta = lstrVM;
                id.Children.Add(newT1);

            }
        }




        public bool IsNumeric(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;
            else
                return value.All(char.IsNumber);
        }

        private void tlistNormativ_NodeCellStyle(object sender, DevExpress.XtraTreeList.GetCustomNodeCellStyleEventArgs e)
        {

            if (e.Column.Name == "treeListColKS" || e.Column.Name == "treeListColID")
            {
                if (e.Node.GetValue(e.Column.AbsoluteIndex) == null)
                {
                    e.Appearance.BackColor = Color.FromArgb(100, 255, 0, 0);
                }

                else if (!(bool)e.Node.GetValue(treeListColNew.AbsoluteIndex))
                {
                    e.Appearance.BackColor = Color.FromArgb(50, 255, 255, 0);
                }

            }

            else if (e.Column.Name == "treeListColNO" || e.Column.Name == "treeListColND" || e.Column.Name == "treeListColJM" || e.Column.Name == "treeListColNew")
            {
                if (e.Node.GetValue(treeListColID.AbsoluteIndex) == null)
                {
                    e.Appearance.BackColor = Color.FromArgb(100, 255, 0, 0);
                }
            }
        }

        private void barButtonItemRefresh_ItemClick(object sender, ItemClickEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            LoadrepositoryItemSearchLookUpEditID();
            foreach (extendedt1 currt1 in mGlobalNormativ)
            {
                RefreshID(currt1);
                currt1.ExistAllMaterialIDS = ProvjeriPostojanjeSifara(currt1);
            }

            LoadrepositoryItemSearchLookUpEditID();
            tlistNormativ.RefreshDataSource();
            tlistNormativ.ExpandAll();

            if (mblnAutoSave) SaveToIvis();
            Cursor = Cursors.Default;

        }

        private void RefreshID(extendedt1 parentt1)
        {
            foreach (extendedt1 currt1 in parentt1.Children)
            {
                if (currt1.t1id == null && currt1.t1ts == "MT")
                {
                    xSiIdents lxSiIdent = new xSiIdents();

                    List<xSiIdent> idList = lxSiIdent.GetSiBySivaSivf("31115", currt1.vrsta, currt1.debljina);
                    if (idList.Count > 0)
                    {
                        currt1.t1id = idList.FirstOrDefault().siid;
                        currt1.ks = idList.FirstOrDefault().siks;
                        currt1.kp = idList.FirstOrDefault().sikp;
                        currt1.no = idList.FirstOrDefault().sino;
                        currt1.nd = idList.FirstOrDefault().sind;
                        currt1.jm = idList.FirstOrDefault().sijm;
                        currt1.exists = true;
                        currt1.ExistAllMaterialIDS = true;

                        xSiIdent imali = identi.Where(x => x.siid == currt1.t1id).FirstOrDefault();


                        repositoryItemSearchLookUpEditID.DataSource = identi;
                        repositoryItemSearchLookUpEditIDView.RefreshData();
                    }
                }

                if (currt1.Children != null)
                    RefreshID(currt1);
            }


        }

        private void LoadrepositoryItemSearchLookUpEditID()
        {
            xSiIdents lxSiIdent = new xSiIdents();
            identi = lxSiIdent.GetIdentsByAP("31115");
        }

        private void repositoryItemSearchLookUpEditID_EditValueChanged(object sender, EventArgs e)
        {
            var searchLookUpEdit = sender as SearchLookUpEdit;
            if (searchLookUpEdit == null)
                return;
            RepositoryItemSearchLookUpEdit riSearchLookUpEdit = searchLookUpEdit.Properties as RepositoryItemSearchLookUpEdit;
            xSiIdent rowByKeyValue = riSearchLookUpEdit.GetRowByKeyValue(searchLookUpEdit.EditValue) as xSiIdent;

            tlistNormativ.SetFocusedRowCellValue("ks", rowByKeyValue.siks);
            tlistNormativ.SetFocusedRowCellValue("kp", rowByKeyValue.sikp);
            tlistNormativ.SetFocusedRowCellValue("id", rowByKeyValue.siid);
            tlistNormativ.SetFocusedRowCellValue("no", rowByKeyValue.sino);
            tlistNormativ.SetFocusedRowCellValue("nd", rowByKeyValue.sind);
            tlistNormativ.SetFocusedRowCellValue("jm", rowByKeyValue.sijm);
        }

        private void barEditItemRefresh_EditValueChanged(object sender, EventArgs e)
        {
            int val;
            int.TryParse(barEditItemRefresh.EditValue.ToString(), out val);
            if (val == 0)
            {
                timerrefreshid.Enabled = false;
                timerProgres.Enabled = false;
            }
            else
            {
                timerrefreshid.Interval = val * 1000 * 60;
                progressBarControlTick.Properties.Maximum = val * 60 - 1;
                timerrefreshid.Enabled = true;
                timerProgres.Enabled = true;
                timerrefreshid.Start();
                progressBarControlTick.EditValue = 0;
                timerProgres.Start();
            }

            Properties.Settings.Default["RefreshTime"] = val;
            Properties.Settings.Default.Save();
        }

        private void timerrefreshid_Tick(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;

            timerrefreshid.Enabled = false;
            timerProgres.Enabled = false;
            LoadrepositoryItemSearchLookUpEditID();
            foreach (extendedt1 currt1 in mGlobalNormativ)
            {
                RefreshID(currt1);
                currt1.ExistAllMaterialIDS = ProvjeriPostojanjeSifara(currt1);
            }
            tlistNormativ.RefreshDataSource();
            tlistNormativ.ExpandAll();
            repositoryItemSearchLookUpEditIDView.RefreshData();
            if (mblnAutoSave) SaveToIvis();
            LoadrepositoryItemSearchLookUpEditID();
            timerrefreshid.Enabled = true;
            timerProgres.Enabled = true;
            timerProgres.Start();
            progressBarControlTick.EditValue = 0;
            progressBarControlTick.Update();
            Cursor = Cursors.Default;

        }

        private void barButtonItemSave_ItemClick(object sender, ItemClickEventArgs e)
        {
            SaveToIvis();
        }

        public void SaveToIvis()
        {

            if (!splashScreenManagerW.IsSplashFormVisible)
                splashScreenManagerW.ShowWaitForm();
            bool timerrefreshidenabled = timerrefreshid.Enabled;
            try
            {
                tlistNormativ.Enabled = false;
                Cursor = Cursors.WaitCursor;
                timerrefreshid.Enabled = false;
            }
            catch { }

            xSiIdents lxSiIdents = new xSiIdents();


           
            foreach (xSiIdent currxSiIdent in identiZaDodati)
            {
                lxSiIdents.SaveNewSi(currxSiIdent); 
                mGlobalNormativ.Where(x => x.t1id == currxSiIdent.siid).ToList().Select(c => { c.exists = true; return c; }).ToList();
            }
            identiZaDodati = new List<xSiIdent>();

         
            foreach (extendedt1 currt1 in mGlobalNormativ)
                SaveNormativ(currt1);

            for (int i = mGlobalNormativ.Count - 1; i >= 0; --i)
                if (mGlobalNormativ[i].ExistAllMaterialIDS && mGlobalNormativ[i].Saved)
                {

                    if (mblnNotifyOK)
                    {
                        string mailText = "<h3 style=\"color: #008000;\">USPJEŠAN IMPORT NORMATIVA IZ PDM U IVIS.</h3>";

                        mailText += "<ul>";
                        mailText += CitajSadrzajNormativa(mGlobalNormativ[i]);
                        mailText += "</ul>";
                        mailText += "<p></p>";
                        mailText += "<p style=\"color: #989898;\"><small><i>Ovo je automatski generisana poruka i nemojte odgovarati na nju.</i></small></p>";
                        mailText += "<p style=\"color: #800000;\">IVIS Software by b2b.ba</p>";

                        SendMail(Properties.Settings.Default["MailOK"].ToString(), mailText, "Uspješan import normativa iz PDM u IVIS");
                    }
                    CleanFolder();
                    mGlobalNormativ.RemoveAt(i);
                }

            if (mblnNotifyNOK && MaterijaliBezSifara.Count > 0)
            {

                string mailText = "<h3 style=\"color: red;\">NEUSPJEŠAN IMPORT NORMATIVA IZ PDM U IVIS.</h3>";
                mailText += "<p>Import normativa iz PDM u IVIS nije moguć, jer ne postoje šifre materijala za sljedeće kombinacije karakteristika:</p>";
                mailText += "<ul>";
                foreach (extendedt1 currt1 in MaterijaliBezSifara)
                    mailText += "<li><b>Vrsta materijala: " + currt1.vrsta + ",&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;debljina: " + currt1.debljina + "</b></li>";
                mailText += "</ul>";
                mailText += "<p></p>";
                mailText += "<p>Nakon unosa prvog identa s navedenim karakteristikama u IVIS, uvoz će se izvršiti automatski.</p>";
                mailText += "<p style=\"color: #989898;\"><small><i>Ovo je automatski generisana poruka i nemojte odgovarati na nju.</i></small></p>";
                mailText += "<p style=\"color: #800000;\">IVIS Software by b2b.ba</p>";

                SendMail(Properties.Settings.Default["MailNOK"].ToString(), mailText, "Neuspješan import normativa iz PDM u IVIS");

                MaterijaliBezSifara = new List<extendedt1>();
            }

            SetDataSource(mGlobalNormativ);


            try
            {
                timerrefreshid.Enabled = timerrefreshidenabled;
                Cursor = Cursors.Default;
                tlistNormativ.Enabled = true;
                splashScreenManagerW.CloseWaitForm();
            }
            catch { }
        }

        private string CitajSadrzajNormativa(extendedt1 currt1)
        {
            string returnvalue = "";

            returnvalue = "<li>" + currt1.t1id + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <strong>" + currt1.no + "</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " + currt1.nd + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " + currt1.t1kl + " " + currt1.jm + "</li>";

            if (currt1.Children != null)
            {
                returnvalue += "<ul>";
                foreach (extendedt1 childt1 in currt1.Children)
                {
                    returnvalue += CitajSadrzajNormativa(childt1);
                }
                returnvalue += "</ul>";
            }

            return returnvalue;
        }

        private void SendMail(string strTO, string strBody, string strSubject)
        {
            try
            {
                string host = Properties.Settings.Default["Host"].ToString();
                Int32 port = 0;
                int.TryParse(Properties.Settings.Default["Port"].ToString(), out port);
                string SenderMail = Properties.Settings.Default["Email"].ToString();
                string Username = Properties.Settings.Default["User"].ToString();
                string Password = Properties.Settings.Default["Pass"].ToString();
                SmtpClient smtp = new SmtpClient(host, port);
                smtp.Credentials = new System.Net.NetworkCredential(Username, Password);
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(SenderMail, "IVIS T8IM");
                mail.Subject = strSubject;
                mail.Body = strBody;
                mail.To.Add(strTO);
                mail.IsBodyHtml = true;
                smtp.Send(mail);
            }
            catch { }
        }

        private bool ProvjeriPostojanjeSifara(extendedt1 currt1)
        {

            bool RetrunValue = true;
            bool tmpValue;
            if (currt1.Children != null)
            {
                foreach (extendedt1 childt1 in currt1.Children)
                {
                    tmpValue = ProvjeriPostojanjeSifara(childt1);

                    if (RetrunValue) RetrunValue = tmpValue;

                }
            }
            else if (currt1.t1ts == "MT")
            {
                RetrunValue = currt1.ExistAllMaterialIDS;
                if (!currt1.ExistAllMaterialIDS)
                    MaterijaliBezSifara.Add(currt1);


            }
            return RetrunValue;
        }

        private void SaveNormativ(extendedt1 currt1)
        {
            tT0P ltT0P = new tT0P();
            if (currt1.t1ts != "MT")
            {
                t0 lt0 = ltT0P.GetT0By(currt1.t1id, currt1.t1ts).T0;
                if (lt0 == null || mblnbarEditItemOverwriteT)
                {
                    lt0 = new t0();
                    lt0.t0id = currt1.t1id;
                    lt0.t0ts = currt1.t1ts;
                    lt0.t0tm = currt1.ks;
                    lt0.t0kl = 1;
                    lt0.t0dr = DateTime.Now;
                    lt0.t0el = 0;
                    lt0.t0er = 0;
                    lt0.t0em = 0;
                    lt0.t0eo = 0;
                    lt0.t0sl = 0;
                    lt0.t0sr = 0;
                    lt0.t0sm = 0;
                    lt0.t0so = 0;
                    lt0.t0ma = 0;
                    lt0.t0st = "N";
                    ltT0P.SaveT0(lt0);
                }
            }
            if (currt1.Children != null)
            {
                tT1P ltT1P = new tT1P();
                int i = 0;
                if (mblnbarEditItemOverwriteT)
                    ltT1P.DeleteT1By(currt1.t1id, currt1.t1ts);
                foreach (extendedt1 xt1child in currt1.Children)
                {
                    t1 newt1 = xt1child.ToT1();
                    if (newt1.t1id != null)
                    {
                        i += 1;
                        newt1.t1rb = i;

                        t1 t1loc = ltT1P.GetT1By(newt1.t1idm, newt1.t1tsm, newt1.t1rb);

                        if (t1loc != null || mblnbarEditItemOverwriteT)
                            ltT1P.SaveT1(newt1);
                        SaveNormativ(xt1child);
                    }

                }
            }

            currt1.Saved = true;
        }

        private void barEditItemAutoSave_EditValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["AutoSave"] = barEditItemAutoSave.EditValue;
            mblnAutoSave = (bool)barEditItemAutoSave.EditValue;
            Properties.Settings.Default.Save();
        }


        private void barEditItemOverwriteT_EditValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["OverwriteT"] = barEditItemOverwriteT.EditValue;
            Properties.Settings.Default.Save();

            mblnbarEditItemOverwriteT = (bool)barEditItemOverwriteT.EditValue;
        }

        private void barButtonItemMailSettings_ItemClick(object sender, ItemClickEventArgs e)
        {
            T8INSettF T8INSettFForm = new T8INSettF();
            T8INSettFForm.ShowDialog();
        }



        private void barEditItemNotifyNOK_EditValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["NotifyNOK"] = barEditItemNotifyNOK.EditValue;
            mblnNotifyNOK = (bool)Properties.Settings.Default["NotifyNOK"];
            Properties.Settings.Default.Save();
        }

        private void barEditItemNotifyOK_EditValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["NotifyOK"] = barEditItemNotifyOK.EditValue;
            mblnNotifyOK = (bool)Properties.Settings.Default["NotifyOK"];
            Properties.Settings.Default.Save();
        }

        private void timerProgres_Tick(object sender, EventArgs e)
        {
            progressBarControlTick.PerformStep();

            if ((int)progressBarControlTick.EditValue == progressBarControlTick.Properties.Maximum)
                barEditItemRefresh_EditValueChanged(null, null);

        }
    }
}
