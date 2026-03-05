using System;
using System.Linq;
using System.Windows;

namespace Kursach.ADialogs
{
    public partial class EventEditDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private int? _eventId;

        public EventEditDialog()
        {
            InitializeComponent();
            TitleText.Text = "Добавление мероприятия";
            LoadEventTypes();
            EventDatePicker.SelectedDate = DateTime.Today;
        }

        public EventEditDialog(int eventId)
        {
            InitializeComponent();
            _eventId = eventId;
            TitleText.Text = "Редактирование мероприятия";
            LoadEventTypes();
            LoadEventData();
        }

        private void LoadEventTypes()
        {
            try
            {
                var types = db.EventTypes
                    .OrderBy(t => t.EventTypeName)
                    .Select(t => new { t.EventTypeID, t.EventTypeName })
                    .ToList();

                EventTypeComboBox.ItemsSource = types;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки типов: {ex.Message}");
            }
        }

        private void LoadEventData()
        {
            try
            {
                if (!_eventId.HasValue) return;

                var ev = db.Events.Find(_eventId.Value);
                if (ev == null) return;

                EventNameTextBox.Text = ev.EventName;
                EventTypeComboBox.SelectedValue = ev.EventTypeID;
                EventDatePicker.SelectedDate = ev.EventDate;

                if (ev.EventTime.HasValue)
                    EventTimeTextBox.Text = ev.EventTime.Value.ToString(@"hh\:mm");

                LocationTextBox.Text = ev.Location;
                OrganizerTextBox.Text = ev.Organizer;
                IsRequiredCheckBox.IsChecked = ev.IsRequired;
                DescriptionTextBox.Text = ev.Description;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(EventNameTextBox.Text))
                {
                    MessageBox.Show("Введите название мероприятия");
                    EventNameTextBox.Focus();
                    return;
                }

                if (EventTypeComboBox.SelectedItem == null)
                {
                    MessageBox.Show("Выберите тип мероприятия");
                    EventTypeComboBox.Focus();
                    return;
                }

                if (!EventDatePicker.SelectedDate.HasValue)
                {
                    MessageBox.Show("Выберите дату мероприятия");
                    EventDatePicker.Focus();
                    return;
                }

                TimeSpan? eventTime = null;
                if (!string.IsNullOrWhiteSpace(EventTimeTextBox.Text))
                {
                    if (TimeSpan.TryParse(EventTimeTextBox.Text, out TimeSpan time))
                        eventTime = time;
                }

                if (_eventId == null)
                {
                    var newEvent = new Events
                    {
                        EventName = EventNameTextBox.Text.Trim(),
                        EventTypeID = (int)EventTypeComboBox.SelectedValue,
                        EventDate = EventDatePicker.SelectedDate,
                        EventTime = eventTime,
                        Location = LocationTextBox.Text?.Trim(),
                        Organizer = OrganizerTextBox.Text?.Trim(),
                        Description = DescriptionTextBox.Text?.Trim(),
                        IsRequired = IsRequiredCheckBox.IsChecked,
                        CreatedAt = DateTime.Now
                    };

                    db.Events.Add(newEvent);
                    db.SaveChanges();

                    MessageBox.Show("Мероприятие успешно добавлено!", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    var ev = db.Events.Find(_eventId.Value);
                    if (ev == null) return;

                    ev.EventName = EventNameTextBox.Text.Trim();
                    ev.EventTypeID = (int)EventTypeComboBox.SelectedValue;
                    ev.EventDate = EventDatePicker.SelectedDate;
                    ev.EventTime = eventTime;
                    ev.Location = LocationTextBox.Text?.Trim();
                    ev.Organizer = OrganizerTextBox.Text?.Trim();
                    ev.Description = DescriptionTextBox.Text?.Trim();
                    ev.IsRequired = IsRequiredCheckBox.IsChecked;

                    db.SaveChanges();

                    MessageBox.Show("Мероприятие успешно обновлено!", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}