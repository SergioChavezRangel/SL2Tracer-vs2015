using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tracer.automation.level2;

namespace TracerTest
{
    public partial class TestApp : Form
    {
        public TestApp()
        {
            InitializeComponent();

            TraceManager logger = new TraceManager();
            
            logger.CRITICAL("start");
        }
    }
}
