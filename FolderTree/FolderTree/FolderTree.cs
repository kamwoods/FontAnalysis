using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms ;
using System.IO ;
using System.ComponentModel;

namespace FolderTree
{
    public class FolderTree : System.Windows.Forms.TreeView, ISupportInitialize
    {
        private string  _root_folder = "";
        private bool    _show_files = true;
        private bool    _in_init = false;

        public FolderTree()
        {
        }

        [Category("Behavior"),
            Description("Gets or sets the root folder of the tree."),
            DefaultValue("C:\\")]
        public string RootFolder
        {
            get{ return _root_folder; }
            set
            {
                _root_folder = value;
                if(!_in_init)
                    InitializeTree();
            }
        }

        [Category("Behavior"),
            Description("Indicates whether files will be seen in list."),
            DefaultValue(true)]
        public bool ShowFiles
        {
            get{ return _show_files; }
            set{ _show_files = value; }
        }

        [Browsable(false)]
        public string SelectedFolder
        {
            get
            {
                if (this.SelectedNode is FolderNode)
                    return ((FolderNode)this.SelectedNode).FolderPath;
                else
                    return "";

            }
        }

        private void LoadTree( FolderNode folder )
        {
            string[] dirs = Directory.GetDirectories(folder.FolderPath);  
  
            foreach( string dir in dirs )
            {
                FolderNode tmp_folder = new FolderNode(dir);
                folder.Nodes.Add(tmp_folder);
                LoadTree( tmp_folder );
            }

            if( _show_files )
            {
                string[] files = Directory.GetFiles(folder.FolderPath);

                foreach( string file in files )
                {
                    FileNode fnode = new FileNode(file);
                    folder.Nodes.Add(fnode);
                }
            }
        }
    
        private void InitializeTree()
        {
            if(!this.DesignMode && _root_folder != "")
            {
                FolderNode rootNode = new FolderNode(_root_folder);
                LoadTree(rootNode);
                this.Nodes.Clear();
                this.Nodes.Add(rootNode);
            }
        }


#region ISupportInitialize Members

        void ISupportInitialize.BeginInit()
        {
            _in_init = true;
        }

        void ISupportInitialize.EndInit()
        {
            if (_root_folder != "")
                InitializeTree();

            _in_init = false;
        }

#endregion
    }

    
    public class FileNode : System.Windows.Forms.TreeNode
    {
        string _file_name = "";
        FileInfo _info;

        public FileNode(string fname)
        {
            _file_name = fname;
            _info = new FileInfo(_file_name);
            base.Text = _info.Name;
            if (_info.Extension.ToLower() == ".exe")
                this.ForeColor = System.Drawing.Color.Red;

        }

        public string FileName
        {
            get { return _file_name; }
            set { _file_name = value; }
        }

        public FileInfo FileNodeInfo
        {
            get { return _info; }
        }
    }

    public class FolderNode : System.Windows.Forms.TreeNode
    {
        string _folder_path = "";
        DirectoryInfo _info;

        public FolderNode(string fname)
        {
            _folder_path = fname;
            _info = new DirectoryInfo(_folder_path);
            base.Text = _info.Name;
        }

        public string FolderPath
        {
            get { return _folder_path; }
            set { _folder_path = value; }
        }

        public DirectoryInfo FolderNodeInfo
        {
            get { return _info; }
        }
    }
        
}
