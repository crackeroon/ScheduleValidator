
namespace ScheduleValidator
{
    partial class LoginForm
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
            this.login_label = new System.Windows.Forms.Label();
            this.name = new System.Windows.Forms.TextBox();
            this.password_label = new System.Windows.Forms.Label();
            this.pass = new System.Windows.Forms.TextBox();
            this.registration = new System.Windows.Forms.Button();
            this.log_in = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // login_label
            // 
            this.login_label.AutoSize = true;
            this.login_label.Location = new System.Drawing.Point(96, 120);
            this.login_label.Name = "login_label";
            this.login_label.Size = new System.Drawing.Size(197, 25);
            this.login_label.TabIndex = 0;
            this.login_label.Text = "Имя пользователя";
            // 
            // name
            // 
            this.name.Location = new System.Drawing.Point(334, 120);
            this.name.Name = "name";
            this.name.Size = new System.Drawing.Size(263, 31);
            this.name.TabIndex = 1;
            // 
            // password_label
            // 
            this.password_label.AutoSize = true;
            this.password_label.Location = new System.Drawing.Point(101, 195);
            this.password_label.Name = "password_label";
            this.password_label.Size = new System.Drawing.Size(86, 25);
            this.password_label.TabIndex = 2;
            this.password_label.Text = "Пароль";
            // 
            // pass
            // 
            this.pass.Location = new System.Drawing.Point(334, 195);
            this.pass.Name = "pass";
            this.pass.PasswordChar = '*';
            this.pass.Size = new System.Drawing.Size(263, 31);
            this.pass.TabIndex = 3;
            // 
            // registration
            // 
            this.registration.Location = new System.Drawing.Point(508, 327);
            this.registration.Name = "registration";
            this.registration.Size = new System.Drawing.Size(180, 47);
            this.registration.TabIndex = 5;
            this.registration.Text = "Регистрация";
            this.registration.UseVisualStyleBackColor = true;
            this.registration.Click += new System.EventHandler(this.registration_Click);
            // 
            // log_in
            // 
            this.log_in.Location = new System.Drawing.Point(101, 327);
            this.log_in.Name = "log_in";
            this.log_in.Size = new System.Drawing.Size(180, 47);
            this.log_in.TabIndex = 4;
            this.log_in.Text = "Войти";
            this.log_in.UseVisualStyleBackColor = true;
            this.log_in.Click += new System.EventHandler(this.log_in_Click);
            // 
            // LoginForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.log_in);
            this.Controls.Add(this.registration);
            this.Controls.Add(this.pass);
            this.Controls.Add(this.password_label);
            this.Controls.Add(this.name);
            this.Controls.Add(this.login_label);
            this.Name = "LoginForm";
            this.Text = "Вход в систему";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label login_label;
        private System.Windows.Forms.TextBox name;
        private System.Windows.Forms.Label password_label;
        private System.Windows.Forms.TextBox pass;
        private System.Windows.Forms.Button registration;
        private System.Windows.Forms.Button log_in;
    }
}