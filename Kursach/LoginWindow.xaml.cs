using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace Kursach
{
    public partial class LoginWindow : Window
    {
        private vsstuEntities db = new vsstuEntities();

        public LoginWindow()
        {
            InitializeComponent();
            LoginTextBox.Focus();
            this.PreviewKeyDown += LoginWindow_PreviewKeyDown;
        }

        private void LoginWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                LoginButton_Click(sender, e);
            }
        }

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string login = LoginTextBox.Text.Trim();
                string password = PasswordBox.Password;

                if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show("Введите логин и пароль", "Ошибка входа",
                                  MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (login == "123" && password == "123")
                {
                    MainWindow mainWindow = new MainWindow();
                    mainWindow.UserRole = "Администратор";
                    mainWindow.UserName = "Администратор";
                    mainWindow.Show();
                    this.Close();
                    return;
                }

                var curator = db.Curators
                    .FirstOrDefault(c => c.Login == login && c.IsActive == true);

                if (curator != null)
                {
                    if (password == "123")
                    {
                        MainWindow mainWindow = new MainWindow();
                        mainWindow.UserRole = "Куратор";
                        mainWindow.UserName = $"{curator.LastName} {curator.FirstName}";
                        mainWindow.UserId = curator.CuratorID;
                        mainWindow.Show();
                        this.Close();
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Неверный пароль", "Ошибка входа",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                        PasswordBox.Clear();
                        PasswordBox.Focus();
                        return;
                    }
                }

                var student = db.Students
                    .FirstOrDefault(s => s.PersonalNumber == login && s.IsActive == true);

                if (student == null)
                {
                    student = db.Students
                        .FirstOrDefault(s => s.StudentCardNumber == login && s.IsActive == true);
                }

                if (student != null)
                {
                    bool isHeadman = student.IsHeadman ?? false;

                    string expectedPassword = "";

                    if (!string.IsNullOrEmpty(student.PersonalNumber) && student.PersonalNumber.Length >= 6)
                    {
                        expectedPassword = student.PersonalNumber.Substring(
                            student.PersonalNumber.Length - 6);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(student.StudentCardNumber))
                        {
                            string cardNumber = student.StudentCardNumber.Replace("СТ-", "").Replace("-", "");
                            if (cardNumber.Length >= 6)
                            {
                                expectedPassword = cardNumber.Substring(cardNumber.Length - 6);
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(expectedPassword) || expectedPassword.Length < 6)
                    {
                        MessageBox.Show($"Не удалось сформировать пароль для студента {student.LastName} {student.FirstName}. Обратитесь к администратору.",
                                      "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"Логин: {login}, Пароль: {expectedPassword}");

                    if (password == expectedPassword)
                    {
                        if (!isHeadman)
                        {
                            MessageBox.Show("Вы не являетесь старостой группы. Доступ запрещен.",
                                          "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        MainWindow mainWindow = new MainWindow();
                        mainWindow.UserRole = "Староста";
                        mainWindow.UserName = $"{student.LastName} {student.FirstName}";
                        mainWindow.UserId = student.StudentID;
                        mainWindow.UserGroupId = student.GroupID;
                        mainWindow.Show();
                        this.Close();
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Неверный пароль", "Ошибка входа",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                        PasswordBox.Clear();
                        PasswordBox.Focus();
                        return;
                    }
                }

                MessageBox.Show("Пользователь с таким логином не найден", "Ошибка входа",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                LoginTextBox.Clear();
                PasswordBox.Clear();
                LoginTextBox.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при входе: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}