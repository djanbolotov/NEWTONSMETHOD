namespace POCKETSEARCHMETHOD
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label19 = new System.Windows.Forms.Label();
            this.ParametrR = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.SearchStep = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.LimitOfTime = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.LimitOfIterations = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.Tolerance = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.InitialApproximation = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Function = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.absError = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.fxminustolerance = new System.Windows.Forms.TextBox();
            this.fxplustolerance = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.ResultX = new System.Windows.Forms.TextBox();
            this.countofiterations = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.elapsedtime = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.fxminusminustolerance = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.fxminusplustolerance = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.fx = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.FunctionGraph = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.validation = new System.Windows.Forms.RichTextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label1.Location = new System.Drawing.Point(385, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(264, 37);
            this.label1.TabIndex = 0;
            this.label1.Text = "Newton\'s Method";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.groupBox1.Controls.Add(this.label19);
            this.groupBox1.Controls.Add(this.ParametrR);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.SearchStep);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.LimitOfTime);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.LimitOfIterations);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.Tolerance);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.InitialApproximation);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.Function);
            this.groupBox1.Location = new System.Drawing.Point(22, 49);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(470, 405);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Input Data:";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(6, 236);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(191, 20);
            this.label19.TabIndex = 17;
            this.label19.Text = "(for constriction area size)";
            // 
            // ParametrR
            // 
            this.ParametrR.Location = new System.Drawing.Point(205, 217);
            this.ParametrR.Name = "ParametrR";
            this.ParametrR.Size = new System.Drawing.Size(259, 26);
            this.ParametrR.TabIndex = 16;
            this.ParametrR.TextChanged += new System.EventHandler(this.ParametrR_TextChanged);
            this.ParametrR.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ParametrR_KeyPress);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(8, 217);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(103, 20);
            this.label18.TabIndex = 15;
            this.label18.Text = "Parameter R:";
            this.label18.UseMnemonic = false;
            // 
            // SearchStep
            // 
            this.SearchStep.Location = new System.Drawing.Point(204, 174);
            this.SearchStep.Name = "SearchStep";
            this.SearchStep.Size = new System.Drawing.Size(260, 26);
            this.SearchStep.TabIndex = 11;
            this.SearchStep.TextChanged += new System.EventHandler(this.SearchStep_TextChanged);
            this.SearchStep.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.SearchStep_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 177);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(51, 20);
            this.label7.TabIndex = 10;
            this.label7.Text = "Delta:";
            this.label7.UseMnemonic = false;
            // 
            // LimitOfTime
            // 
            this.LimitOfTime.Location = new System.Drawing.Point(205, 259);
            this.LimitOfTime.Name = "LimitOfTime";
            this.LimitOfTime.Size = new System.Drawing.Size(259, 26);
            this.LimitOfTime.TabIndex = 9;
            this.LimitOfTime.TextChanged += new System.EventHandler(this.LimitOfTime_TextChanged);
            this.LimitOfTime.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.LimitOfTime_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 262);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(137, 20);
            this.label6.TabIndex = 8;
            this.label6.Text = "Limit of time (sec):";
            // 
            // LimitOfIterations
            // 
            this.LimitOfIterations.Location = new System.Drawing.Point(205, 302);
            this.LimitOfIterations.Name = "LimitOfIterations";
            this.LimitOfIterations.Size = new System.Drawing.Size(259, 26);
            this.LimitOfIterations.TabIndex = 7;
            this.LimitOfIterations.TextChanged += new System.EventHandler(this.LimitOfIterations_TextChanged);
            this.LimitOfIterations.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.LimitOfIterations_KeyPress);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 305);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(133, 20);
            this.label5.TabIndex = 6;
            this.label5.Text = "Limit of iterations:";
            // 
            // Tolerance
            // 
            this.Tolerance.Location = new System.Drawing.Point(204, 131);
            this.Tolerance.Name = "Tolerance";
            this.Tolerance.Size = new System.Drawing.Size(260, 26);
            this.Tolerance.TabIndex = 5;
            this.Tolerance.TextChanged += new System.EventHandler(this.Tolerance_TextChanged);
            this.Tolerance.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Tolerance_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 134);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 20);
            this.label4.TabIndex = 4;
            this.label4.Text = "Tolerance:";
            // 
            // InitialApproximation
            // 
            this.InitialApproximation.Location = new System.Drawing.Point(204, 89);
            this.InitialApproximation.Name = "InitialApproximation";
            this.InitialApproximation.Size = new System.Drawing.Size(260, 26);
            this.InitialApproximation.TabIndex = 3;
            this.InitialApproximation.TextChanged += new System.EventHandler(this.InitialApproximation_TextChanged);
            this.InitialApproximation.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.InitialApproximation_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 95);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(173, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "Initial approximation х0:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "Function F(x):";
            // 
            // Function
            // 
            this.Function.FormattingEnabled = true;
            this.Function.Items.AddRange(new object[] {
            "(x^2-1)^2",
            "x^2-4*sin(x)",
            "x^3-5*x^2+x+5",
            "(x-2)^2-ln(x)",
            "(x-2)^2-log(x)"});
            this.Function.Location = new System.Drawing.Point(204, 45);
            this.Function.Name = "Function";
            this.Function.Size = new System.Drawing.Size(260, 28);
            this.Function.TabIndex = 2;
            this.Function.Text = "(x^2-1)^2";
            this.Function.TextChanged += new System.EventHandler(this.Function_TextChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.groupBox2.Controls.Add(this.absError);
            this.groupBox2.Controls.Add(this.label21);
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.fxminustolerance);
            this.groupBox2.Controls.Add(this.fxplustolerance);
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Controls.Add(this.ResultX);
            this.groupBox2.Controls.Add(this.countofiterations);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.elapsedtime);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.fxminusminustolerance);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.fxminusplustolerance);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.fx);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Location = new System.Drawing.Point(538, 49);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(508, 417);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Output Data:";
            // 
            // absError
            // 
            this.absError.Enabled = false;
            this.absError.Location = new System.Drawing.Point(193, 360);
            this.absError.Name = "absError";
            this.absError.Size = new System.Drawing.Size(309, 26);
            this.absError.TabIndex = 20;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(6, 363);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(72, 20);
            this.label21.TabIndex = 19;
            this.label21.Text = "RelError:";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 161);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(129, 20);
            this.label15.TabIndex = 16;
            this.label15.Text = "F(x* - Tolerance):";
            this.label15.UseMnemonic = false;
            // 
            // fxminustolerance
            // 
            this.fxminustolerance.Enabled = false;
            this.fxminustolerance.Location = new System.Drawing.Point(193, 158);
            this.fxminustolerance.Name = "fxminustolerance";
            this.fxminustolerance.Size = new System.Drawing.Size(309, 26);
            this.fxminustolerance.TabIndex = 15;
            // 
            // fxplustolerance
            // 
            this.fxplustolerance.Enabled = false;
            this.fxplustolerance.Location = new System.Drawing.Point(193, 119);
            this.fxplustolerance.Name = "fxplustolerance";
            this.fxplustolerance.Size = new System.Drawing.Size(309, 26);
            this.fxplustolerance.TabIndex = 14;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(6, 122);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(133, 20);
            this.label16.TabIndex = 13;
            this.label16.Text = "F(x* + Tolerance):";
            // 
            // ResultX
            // 
            this.ResultX.Enabled = false;
            this.ResultX.Location = new System.Drawing.Point(193, 42);
            this.ResultX.Name = "ResultX";
            this.ResultX.Size = new System.Drawing.Size(309, 26);
            this.ResultX.TabIndex = 12;
            // 
            // countofiterations
            // 
            this.countofiterations.Enabled = false;
            this.countofiterations.Location = new System.Drawing.Point(193, 319);
            this.countofiterations.Name = "countofiterations";
            this.countofiterations.Size = new System.Drawing.Size(309, 26);
            this.countofiterations.TabIndex = 11;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 240);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(175, 20);
            this.label8.TabIndex = 10;
            this.label8.Text = "F(x*) - F(x* - Tolerance):";
            this.label8.UseMnemonic = false;
            // 
            // elapsedtime
            // 
            this.elapsedtime.Enabled = false;
            this.elapsedtime.Location = new System.Drawing.Point(193, 278);
            this.elapsedtime.Name = "elapsedtime";
            this.elapsedtime.Size = new System.Drawing.Size(309, 26);
            this.elapsedtime.TabIndex = 9;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 281);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(105, 20);
            this.label9.TabIndex = 8;
            this.label9.Text = "Elapsed time:";
            // 
            // fxminusminustolerance
            // 
            this.fxminusminustolerance.Enabled = false;
            this.fxminusminustolerance.Location = new System.Drawing.Point(193, 237);
            this.fxminusminustolerance.Name = "fxminusminustolerance";
            this.fxminusminustolerance.Size = new System.Drawing.Size(309, 26);
            this.fxminusminustolerance.TabIndex = 7;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 322);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(143, 20);
            this.label10.TabIndex = 6;
            this.label10.Text = "Count of iterations:";
            // 
            // fxminusplustolerance
            // 
            this.fxminusplustolerance.Enabled = false;
            this.fxminusplustolerance.Location = new System.Drawing.Point(193, 199);
            this.fxminusplustolerance.Name = "fxminusplustolerance";
            this.fxminusplustolerance.Size = new System.Drawing.Size(309, 26);
            this.fxminusplustolerance.TabIndex = 5;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 202);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(179, 20);
            this.label11.TabIndex = 4;
            this.label11.Text = "F(x*) - F(x* + Tolerance):";
            // 
            // fx
            // 
            this.fx.Enabled = false;
            this.fx.Location = new System.Drawing.Point(193, 80);
            this.fx.Name = "fx";
            this.fx.Size = new System.Drawing.Size(309, 26);
            this.fx.TabIndex = 3;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 45);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(86, 20);
            this.label12.TabIndex = 2;
            this.label12.Text = "Result (х*):";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 83);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(46, 20);
            this.label13.TabIndex = 2;
            this.label13.Text = "F(x*):";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.button1.Location = new System.Drawing.Point(20, 513);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 130);
            this.button1.TabIndex = 13;
            this.button1.Text = "Run Method";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.button2.Location = new System.Drawing.Point(335, 588);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(155, 55);
            this.button2.TabIndex = 14;
            this.button2.Text = "Set \'a\' like \'x0\'";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.button3.Location = new System.Drawing.Point(130, 513);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(184, 57);
            this.button3.TabIndex = 15;
            this.button3.Text = "Clear";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // FunctionGraph
            // 
            this.FunctionGraph.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.FunctionGraph.Location = new System.Drawing.Point(130, 588);
            this.FunctionGraph.Name = "FunctionGraph";
            this.FunctionGraph.Size = new System.Drawing.Size(184, 55);
            this.FunctionGraph.TabIndex = 16;
            this.FunctionGraph.Text = "Show function graph";
            this.FunctionGraph.UseVisualStyleBackColor = false;
            this.FunctionGraph.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.button5.Location = new System.Drawing.Point(335, 513);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(155, 57);
            this.button5.TabIndex = 17;
            this.button5.Text = "User manual";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(534, 490);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(83, 20);
            this.label17.TabIndex = 17;
            this.label17.Text = "Validation:";
            // 
            // validation
            // 
            this.validation.Enabled = false;
            this.validation.Location = new System.Drawing.Point(538, 513);
            this.validation.Name = "validation";
            this.validation.Size = new System.Drawing.Size(508, 130);
            this.validation.TabIndex = 18;
            this.validation.Text = "";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(22, 460);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(470, 32);
            this.progressBar1.TabIndex = 19;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1058, 679);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.validation);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.FunctionGraph);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "NM_Zhanbolot_uulu_Askabek";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox SearchStep;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox LimitOfTime;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox LimitOfIterations;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Tolerance;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox InitialApproximation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox Function;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox countofiterations;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox elapsedtime;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox fxminusminustolerance;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox fxminusplustolerance;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox fx;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox ResultX;
        private System.Windows.Forms.TextBox ParametrR;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox fxminustolerance;
        private System.Windows.Forms.TextBox fxplustolerance;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button FunctionGraph;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.RichTextBox validation;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox absError;
        private System.Windows.Forms.Label label21;
    }
}

