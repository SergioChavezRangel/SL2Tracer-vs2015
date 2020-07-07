using System;
using System.Collections;
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
            
            logger.CRITICAL("critical");
            logger.EXCEPTION("exception");
            logger.WARNING("warning");
            logger.INFO("info");
            logger.DEBUG("debug");
            logger.VERBOSE("verbose");

            Hashtable ht = new Hashtable();
            Hashtable ht2 = new Hashtable();
            ht2.Add("SecondOne", 100);
            ht2.Add("SecondTwo", 200);
            ht.Add("One", 1);
            ht.Add("Two", 2);
            ht.Add("Three", "Three");
            ht.Add("Four", ht2);
            logger.INFO(ht);


            Dictionary<string, Hashtable> dict = new Dictionary<string, Hashtable>();
            dict.Add("firstDictEntry", ht);
            dict.Add("nextDictEntry", ht2);
            logger.INFO(dict);

        }
    }
}
