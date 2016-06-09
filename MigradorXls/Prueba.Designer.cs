namespace MigradorXls
{
    partial class Prueba
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Excel (.xls)");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Migracion por Formato", new System.Windows.Forms.TreeNode[] {
            treeNode1});
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("A2");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Migracion por Base de Datos", new System.Windows.Forms.TreeNode[] {
            treeNode3});
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Reportes");
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Node6");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Prueba));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Left;
            this.treeView1.ImageIndex = 0;
            this.treeView1.ImageList = this.imageList1;
            this.treeView1.Location = new System.Drawing.Point(0, 0);
            this.treeView1.Name = "treeView1";
            treeNode1.Name = "NodeExcel";
            treeNode1.Text = "Excel (.xls)";
            treeNode2.ImageIndex = 0;
            treeNode2.Name = "Node0";
            treeNode2.Text = "Migracion por Formato";
            treeNode2.ToolTipText = "Migracion a travez de los formatos pre establecidos";
            treeNode3.Name = "Node4";
            treeNode3.Text = "A2";
            treeNode3.ToolTipText = "Migracion desde base de datos de A2";
            treeNode4.ImageIndex = 1;
            treeNode4.Name = "Node3";
            treeNode4.Text = "Migracion por Base de Datos";
            treeNode4.ToolTipText = "Migracion directamente de la base de datos de varios sistemas administrativos (a2" +
    ", premium soft,etc)";
            treeNode5.ImageIndex = 2;
            treeNode5.Name = "Node5";
            treeNode5.Text = "Reportes";
            treeNode6.ImageIndex = 3;
            treeNode6.Name = "Node6";
            treeNode6.Text = "Node6";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode2,
            treeNode4,
            treeNode5,
            treeNode6});
            this.treeView1.SelectedImageIndex = 0;
            this.treeView1.Size = new System.Drawing.Size(186, 562);
            this.treeView1.TabIndex = 0;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "migrar xls.png");
            this.imageList1.Images.SetKeyName(1, "migrar bases de datos.png");
            this.imageList1.Images.SetKeyName(2, "reportes.png");
            this.imageList1.Images.SetKeyName(3, "salir.png");
            // 
            // Prueba
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.treeView1);
            this.IsMdiContainer = true;
            this.Name = "Prueba";
            this.Text = "Prueba";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ImageList imageList1;
    }
}