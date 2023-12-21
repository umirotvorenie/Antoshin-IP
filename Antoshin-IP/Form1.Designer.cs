namespace Antoshin_IP
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            textBox1 = new TextBox();
            textBox2 = new TextBox();
            textBox3 = new TextBox();
            textBox4 = new TextBox();
            textBox5 = new TextBox();
            textBox6 = new TextBox();
            label1 = new Label();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            label7 = new Label();
            label8 = new Label();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(463, 40);
            button1.Name = "button1";
            button1.Size = new Size(116, 48);
            button1.TabIndex = 0;
            button1.Text = "Рассчитать";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(266, 49);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(55, 23);
            textBox1.TabIndex = 1;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(327, 49);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(55, 23);
            textBox2.TabIndex = 2;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(388, 49);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(55, 23);
            textBox3.TabIndex = 3;
            // 
            // textBox4
            // 
            textBox4.Location = new Point(266, 96);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(55, 23);
            textBox4.TabIndex = 4;
            // 
            // textBox5
            // 
            textBox5.Location = new Point(327, 96);
            textBox5.Name = "textBox5";
            textBox5.Size = new Size(55, 23);
            textBox5.TabIndex = 5;
            // 
            // textBox6
            // 
            textBox6.Location = new Point(388, 96);
            textBox6.Name = "textBox6";
            textBox6.Size = new Size(55, 23);
            textBox6.TabIndex = 6;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 14.25F, FontStyle.Regular, GraphicsUnit.Point);
            label1.Location = new Point(25, 47);
            label1.Name = "label1";
            label1.Size = new Size(231, 25);
            label1.TabIndex = 7;
            label1.Text = "Линия коэффициентов 1 ";
            // 
            // button2
            // 
            button2.Location = new Point(463, 94);
            button2.Name = "button2";
            button2.Size = new Size(116, 48);
            button2.TabIndex = 13;
            button2.Text = "Показать в Word";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(463, 148);
            button3.Name = "button3";
            button3.Size = new Size(116, 48);
            button3.TabIndex = 14;
            button3.Text = "Показать в Excel";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(463, 201);
            button4.Name = "button4";
            button4.Size = new Size(116, 48);
            button4.TabIndex = 15;
            button4.Text = "Показать в PDF";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Point);
            label7.Location = new Point(37, 171);
            label7.Name = "label7";
            label7.Size = new Size(0, 25);
            label7.TabIndex = 16;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Segoe UI", 14.25F, FontStyle.Regular, GraphicsUnit.Point);
            label8.Location = new Point(25, 94);
            label8.Name = "label8";
            label8.Size = new Size(231, 25);
            label8.TabIndex = 17;
            label8.Text = "Линия коэффициентов 2 ";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(591, 450);
            Controls.Add(label8);
            Controls.Add(label7);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(label1);
            Controls.Add(textBox6);
            Controls.Add(textBox5);
            Controls.Add(textBox4);
            Controls.Add(textBox3);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Найти острый угол между прямыми заданными тремя различными способами";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private TextBox textBox1;
        private TextBox textBox2;
        private TextBox textBox3;
        private TextBox textBox4;
        private TextBox textBox5;
        private TextBox textBox6;
        private Label label1;
        private Button button2;
        private Button button3;
        private Button button4;
        private Label label7;
        private Label label8;
    }
}