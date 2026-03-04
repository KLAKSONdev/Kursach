using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Kursach.ADialogs;

namespace Kursach.AWindows
{
    public partial class EventsWindow : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private bool _isAdmin;
        private int? _groupId;
        private List<EventListItem> allEvents = new List<EventListItem>();

        public EventsWindow(bool isAdminMode, int? groupId = null)
        {
            InitializeComponent();
            _isAdmin = isAdminMode;
            _groupId = groupId;

            // Настраиваем видимость кнопок
            AddEventBtn.Visibility = _isAdmin ? Visibility.Visible : Visibility.Collapsed;
            EditEventBtn.Visibility = _isAdmin ? Visibility.Visible : Visibility.Collapsed;
            DeleteEventBtn.Visibility = _isAdmin ? Visibility.Visible : Visibility.Collapsed;

            LoadEventTypes();
            LoadEvents();
        }

        private void LoadEventTypes()
        {
            try
            {
                var types = db.EventTypes.Select(t => t.EventTypeName).ToList();
                TypeFilter.Items.Clear();
                TypeFilter.Items.Add(new ComboBoxItem { Content = "Все типы", IsSelected = true });
                foreach (var type in types)
                {
                    TypeFilter.Items.Add(new ComboBoxItem { Content = type });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки типов: {ex.Message}");
            }
        }

        private void LoadEvents()
        {
            try
            {
                var query = db.Events.Select(e => new EventListItem
                {
                    EventID = e.EventID,
                    EventName = e.EventName,
                    EventType = e.EventTypes.EventTypeName,
                    EventDate = e.EventDate ?? DateTime.Now,
                    EventTime = e.EventTime,
                    Location = e.Location,
                    Organizer = e.Organizer,
                    ParticipantsCount = e.EventParticipation.Count
                });

                allEvents = query.OrderByDescending(e => e.EventDate).ToList();
                EventsGrid.ItemsSource = allEvents;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки мероприятий: {ex.Message}");
            }
        }

        private void Filter_Changed(object sender, EventArgs e)
        {
            try
            {
                var filtered = allEvents.AsEnumerable();

                var typeItem = TypeFilter.SelectedItem as ComboBoxItem;
                if (typeItem != null && typeItem.Content.ToString() != "Все типы")
                {
                    filtered = filtered.Where(ev => ev.EventType == typeItem.Content.ToString());
                }

                if (DateFilter.SelectedDate.HasValue)
                {
                    filtered = filtered.Where(ev => ev.EventDate.Date == DateFilter.SelectedDate.Value.Date);
                }

                EventsGrid.ItemsSource = filtered.ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
            }
        }

        private void AddEvent_Click(object sender, RoutedEventArgs e)
        {
            if (!_isAdmin) return;

            var dialog = new ADialogs.EventEditDialog();
            dialog.Owner = this;

            if (dialog.ShowDialog() == true)
            {
                LoadEvents();
            }
        }

        private void EditEvent_Click(object sender, RoutedEventArgs e)
        {
            if (!_isAdmin || EventsGrid.SelectedItem == null) return;

            var selected = EventsGrid.SelectedItem as EventListItem;
            if (selected == null) return;

            var dialog = new ADialogs.EventEditDialog(selected.EventID);
            dialog.Owner = this;

            if (dialog.ShowDialog() == true)
            {
                LoadEvents();
            }
        }

        private void DeleteEvent_Click(object sender, RoutedEventArgs e)
        {
            if (!_isAdmin || EventsGrid.SelectedItem == null) return;

            var selected = EventsGrid.SelectedItem as EventListItem;
            if (selected == null) return;

            var result = MessageBox.Show($"Удалить мероприятие '{selected.EventName}'?\n\n" +
                                         $"Участников: {selected.ParticipantsCount}\n" +
                                         $"Все связанные записи об участии будут также удалены.",
                "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var eventToDelete = db.Events.Find(selected.EventID);
                    if (eventToDelete != null)
                    {
                        db.Events.Remove(eventToDelete); // Удалятся и все Participation из-за каскадного удаления
                        db.SaveChanges();
                        LoadEvents();
                        MessageBox.Show("Мероприятие удалено", "Успех");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления: {ex.Message}");
                }
            }
        }

        private void Participate_Click(object sender, RoutedEventArgs e)
        {
            if (EventsGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите мероприятие");
                return;
            }

            var selected = EventsGrid.SelectedItem as EventListItem;
            var dialog = new ParticipationWindow(selected.EventID, _groupId);
            dialog.Owner = this;

            if (dialog.ShowDialog() == true)
            {
                LoadEvents(); // обновляем количество участников
            }
        }

        private void EventsGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_isAdmin)
                EditEvent_Click(sender, e);
            else
                Participate_Click(sender, e);
        }
    }

    // Модели данных
    public class EventListItem
    {
        public int EventID { get; set; }
        public string EventName { get; set; }
        public string EventType { get; set; }
        public DateTime EventDate { get; set; }
        public TimeSpan? EventTime { get; set; }
        public string Location { get; set; }
        public string Organizer { get; set; }
        public int ParticipantsCount { get; set; }
    }
}