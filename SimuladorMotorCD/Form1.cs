using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace SimuladorMotorCD
{

    public partial class Form1 : Form
    {
        public const int maxNeuronas = 10;
        public const int inLayerNeurons = 2;
        public const int hidLayerNeurons = 9;
        public const int outLayerNeurons = 1;
        public const int patterns = 25;
        public const double learningRate = 1;
        public const double allowedError = 0.01;
        public const double ejecucionesLim = 100000000;

        public double[] trainingVector1 = new double[patterns];
        public double[] trainingVector2 = new double[patterns];
        public double[] targetVector1 = new double[patterns];

        public double ejecuciones;
        public int n = 0;
        public bool stopTraining = false;
        public bool[] wellTrained = new bool[patterns];

        public struct neurona
        {
            public double bias;
            public double output;
            public double outputDer;
            public double error;
            public double[] weightCorrection;
            public double biasCorrection;
            public double[] w;
            //public double[] w = new double[maxNeuronas];
        }

        public struct layer
        {
            public neurona[] neurons;
            //public neurona[] neurons = new neurona[maxNeuronas];
        }

        layer inLayer = new layer();        //Capa de entrada de red neuronal
        layer hidLayer = new layer();       //Capa oculta de red neuronal
        layer outLayer = new layer();       //Capa de salida de red neuronal
        
        double k1, k2, k3, k4;              //Valores calculados por RungeKutta
        double[] i = new double[1000];      //Corriente de armadura de motor
        double []Eb = new double[1000];     //Fuerza contraelectromotriz
        double []Tt = new double[1000];     //Par del motor
        double[] w = new double[1000];      //Velocidad angular del motor
        int count;                          //numero de ejecucion
        double CO, COSat;                   //Salida de control, salida de control saturada
        double error;                       //Error (sp - vp)
        double errorAnt = 0;                //Error anterior
        double derivadaError = 0;           //Derivada del error

        public Form1()
        {
            InitializeComponent();
            numericUpDownEa.Value = 0;
            numericUpDownRa.Value = (decimal)0.316;
            numericUpDownLa.Value = (decimal)0.082;
            numericUpDownKt.Value = (decimal)30.2;
            numericUpDownKb.Value = (decimal)317;
            numericUpDownJ.Value = (decimal)0.139;
            numericUpDownB.Value = (decimal)0;
            numericUpDownH.Value = (decimal)0.0005;

            inLayer.neurons = new neurona[maxNeuronas];
            hidLayer.neurons = new neurona[maxNeuronas];
            outLayer.neurons = new neurona[maxNeuronas];

            for (int i = 0; i < maxNeuronas; i++)
            {
                inLayer.neurons[i].w = new double[maxNeuronas];
                hidLayer.neurons[i].w = new double[maxNeuronas];
                outLayer.neurons[i].w = new double[maxNeuronas];

                inLayer.neurons[i].weightCorrection = new double[maxNeuronas];
                hidLayer.neurons[i].weightCorrection = new double[maxNeuronas];
                outLayer.neurons[i].weightCorrection = new double[maxNeuronas];
            }

            iniciaValores();
        }

        void iniciaValores()
        {
            Random random = new Random();

            //Inicializacion de neuronas de capa de entrada
            //Pesos y salidas aleatorias, bias en cero ya que estas neuronas no lo usan
            for (int i = 0; i < inLayerNeurons; i++)
            {
                inLayer.neurons[i].bias = 0;
                inLayer.neurons[i].output = random.NextDouble();
                for (int j = 0; j < hidLayerNeurons; j++)
                {
                    inLayer.neurons[i].w[j] = random.NextDouble();
                }
            }

            //Inicilizacion de neuronas de capa oculta
            //Pesos y polarizaciones aleatorias, salidas en cero al ser calculadas posteriormente
            for (int j = 0; j < hidLayerNeurons; j++)
            {
                hidLayer.neurons[j].bias = random.NextDouble();
                hidLayer.neurons[j].output = 0;
                for (int k = 0; k < outLayerNeurons; k++)
                {
                    hidLayer.neurons[j].w[k] = random.NextDouble();
                }
            }

            //Inicilizacion de neuronas de capa de salida
            //Polarizaciones aleatorias, salidas en cero al ser calculadas posteriormente, pesos no se inician al no ser necesarios
            for (int k = 0; k < outLayerNeurons; k++)
            {
                outLayer.neurons[k].bias = random.NextDouble();
                outLayer.neurons[k].output = 0;
            }
        }

        void calculaSalidas()
        {
            for (int j = 0; j < hidLayerNeurons; j++)
            {
                hidLayer.neurons[j].output = 0;
                for (int i = 0; i < inLayerNeurons; i++)
                {
                    hidLayer.neurons[j].output += inLayer.neurons[i].output * inLayer.neurons[i].w[j];
                }
                hidLayer.neurons[j].output += hidLayer.neurons[j].bias;
                hidLayer.neurons[j].output = 1 / (1 + Math.Exp(-hidLayer.neurons[j].output));
            }

            for (int k = 0; k < outLayerNeurons; k++)
            {
                outLayer.neurons[k].output = 0;
                for (int j = 0; j < hidLayerNeurons; j++)
                {
                    outLayer.neurons[k].output += hidLayer.neurons[j].output * hidLayer.neurons[j].w[k];
                }
                outLayer.neurons[k].output += outLayer.neurons[k].bias;
                outLayer.neurons[k].output = 1 / (1 + Math.Exp(-outLayer.neurons[k].output));
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            i[0] = 0;
            Eb[0] = 0;
            Tt[0] = 0;
            w[0] = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (count = 1; count <= 1000; count++)
            {
                calculos();

                if ((count - 1) % 10 == 0)
                    perceptron();
            }
        }

        private void perceptron ()
        { 
            double sp = (double)numericUpDownSp.Value * (2 * Math.PI) / 60;
            double pv = w[999];

            error = (sp - pv) / (24000 * (2 * Math.PI) / 60) + 0.5; // Calcula error con respecto al doble del error maximo y agrega ofset de 0.5
            derivadaError = (error - errorAnt)/2 + 0.5;

            inLayer.neurons[0].output = error;
            inLayer.neurons[1].output = derivadaError;

            calculaSalidas();

            CO += (outLayer.neurons[0].output - 0.5) * 48;

            if (CO > 24)
                COSat = 24;
            else if (CO < 0)
                COSat = 0;
            else
                COSat = CO;

            numericUpDownEa.Value = (decimal)COSat;

            CO = COSat;
            errorAnt = error;
        }

        private void numericUpDownH_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void calculos()
        {
            double iAct, iSig;
            double wAct, wSig;
            double x, y;
            
            //Lectura de datos de numericUpDown
            double h = (double)(numericUpDownH.Value);
            double La = (double)numericUpDownLa.Value / 1000;
            double Ea = (double)numericUpDownEa.Value;
            double Ra = (double)numericUpDownRa.Value;
            double kt = (double)numericUpDownKt.Value / 1000;
            double kb = (double)(numericUpDownKb.Value) * (double)(1.0 / 60.0) * (double)(2 * Math.PI);
            kb = 1 / kb;
            double J = (double)numericUpDownJ.Value / 1000;
            double B = (double)numericUpDownB.Value;
            double load = (double)numericUpDownLoad.Value;

            //Calculo de corriente por runge kutta
            iAct = i[999];
            k1 = h * (1 / La) * (Ea - (Ra * i[999]) - Eb[999]);
            k2 = h * (1 / La) * (Ea - (Ra * (i[999] + k1 / 2)) - Eb[999]);
            k3 = h * (1 / La) * (Ea - (Ra * (i[999] + k2 / 2)) - Eb[999]);
            k4 = h * (1 / La) * (Ea - (Ra * (i[999] + k3)) - Eb[999]);
            iSig = iAct + (1.0/6.0)*(k1 + (2*k2) + (2*k3) + k4);

            //calculo de torque de motor
            Tt[999] = iAct * kt;

            //calculo de velocidad angular por runge kutta
            wAct = w[999-1];
            k1 = h * (1 / J) * (Tt[999] - (B * w[999]) - load);
            k2 = h * (1 / J) * (Tt[999] - (B * (w[999] + k1 / 2)) - load);
            k3 = h * (1 / J) * (Tt[999] - (B * (w[999] + k2 / 2)) - load);
            k4 = h * (1 / J) * (Tt[999] - (B * (w[999] + k3)) - load);
            wSig = wAct + (1.0 / 6.0) * (k1 + (2 * k2) + (2 * k3) + k4);
            w[999] = wSig;

            //Calculo de FEM
            Eb[999] = (kb * wAct);

            for (int j = 0; j < 999; j++)
            {
                i[j] = i[j + 1];
                w[j] = w[j + 1];
            }
            i[999] = iSig;
            w[999] = wSig;

            impresionResultados();

            /*listBox1.Items.Add(h*count
                                + "\t" + Ea + "V"
                                + "\t" + iAct + "A"
                                + "\t" + Ra + "ohm"
                                + "\t" + La + "mH"
                                + "\t" + Eb[999] + "V"
                                + "\t" + Tt[999] + "Nm"
                                + "\t" + kt + "Nm/A"
                                + "\t" + kb + "V/rad/s"
                                + "\t" + wAct + "rad/s" 
                                );
             */
 
            listBox1.Items.Add(h*count
                                + "\t" + Ea + "V"
                                + "\t \t" + error
                                + "\t \t" + CO);
        }

        private void impresionResultados ()
        {
            double corriente;
            double velAngular;
            double setPoint;
            double t;

            chart1.Series["Corriente (A)"].Points.Clear();
            chart2.Series["Velocidad Angular (RPM)"].Points.Clear();
            chart2.Series["SetPoint"].Points.Clear();

            setPoint = (double)numericUpDownSp.Value;

            for (int k = 0; k < 1000; k++)
            {
                t = (double)(numericUpDownH.Value * (k-999+count));
                corriente = i[k];
                velAngular = w[k] * 60 / (2 * Math.PI);
                
                chart1.Series["Corriente (A)"].Points.AddXY(t, corriente);
                chart2.Series["Velocidad Angular (RPM)"].Points.AddXY(t, velAngular);
                chart2.Series["SetPoint"].Points.AddXY(t, setPoint);
            }
        }

        private void buttonRestart_Click(object sender, EventArgs e)
        {
            count = 0;
            chart1.Series["Corriente (A)"].Points.Clear();
            chart2.Series["Velocidad Angular (RPM)"].Points.Clear();
            chart2.Series["SetPoint"].Points.Clear();

            listBox1.Items.Clear();

            numericUpDownEa.Value = 0;

            for (int k = 0; k < 1000; k++)
            {
                i[k] = 0;
                w[k] = 0;
                Eb[k] = 0;
                Tt[k] = 0;
            }

            errorAnt = 0;
            derivadaError = 0;
            CO = 0;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //Importar patrones desde tabla de excel
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls";
            dlg.ShowDialog();

            try
            {
                string fileName = dlg.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Excel.Worksheet worksheet2;
                Excel.Worksheet worksheet3;
                Excel.Range range;
                workbook = excelApp.Workbooks.Open(fileName);

                int fila, columna;

                worksheet = (Excel.Worksheet)workbook.Sheets["Pesos1"];

                range = worksheet.UsedRange;
                for (fila = 0; fila < inLayerNeurons; fila++)
                {
                    for (columna = 0; columna < hidLayerNeurons; columna++)
                    {
                        inLayer.neurons[fila].w[columna] = Convert.ToDouble((range.Cells[fila + 1, columna + 1] as Excel.Range).Value);
                    }
                }

                worksheet2 = (Excel.Worksheet)workbook.Sheets["Pesos2"];
                worksheet2.Select();
                range = worksheet2.UsedRange;
                for (fila = 0; fila < hidLayerNeurons; fila++)
                {
                    for (columna = 0; columna < outLayerNeurons; columna++)
                    {
                        hidLayer.neurons[fila].w[columna] = Convert.ToDouble((range.Cells[fila + 1, columna + 1] as Excel.Range).Value);
                    }
                }

                worksheet3 = (Excel.Worksheet)workbook.Sheets["Polarizaciones"];
                worksheet3.Select();
                range = worksheet3.UsedRange;
                for (fila = 0; fila < hidLayerNeurons; fila++)
                {
                    hidLayer.neurons[fila].bias = Convert.ToDouble((range.Cells[fila + 1, 1] as Excel.Range).Value);
                }
                for (fila = 0; fila < outLayerNeurons; fila++)
                {
                    outLayer.neurons[fila].bias = Convert.ToDouble((range.Cells[fila + 1, 2] as Excel.Range).Value);
                }
            }

            catch (Exception ex)
            { }
        }
    }
}