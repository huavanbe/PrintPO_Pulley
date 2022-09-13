using System;
using System.Windows;
using System.Windows.Controls;

namespace CustomControl
{
    public class CustomGridPanel : Grid
    {
        /// <summary>
        /// Grid animation actions enum
        /// </summary>
        private enum Action
        {
            Open,
            Close
        }

        #region Variables
        private const double DEFAULT_SPEED = 5;
        private System.Windows.Threading.DispatcherTimer dispatcherTimer;
        private double speed;
        private double exspectedMinWidth;
        private double exspectedMaxWidth;

        #endregion

        public static event PropertyChangedCallback IsOpenChanged;

        /// <summary>
        /// Gets or sets a value indicating whether this instance is open.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is open; otherwise, <c>false</c>.
        /// </value>
        public bool IsOpen
        {
            get { return (bool)GetValue(IsOpenProperty); }
            set { SetValue(IsOpenProperty, value); }
        }

        public static readonly DependencyProperty IsOpenProperty =
            DependencyProperty.Register("IsOpen", typeof(bool), typeof(CustomGridPanel),
                new FrameworkPropertyMetadata(true, FrameworkPropertyMetadataOptions.AffectsRender, IsOpenChangedAction));

        /// <summary>
        /// Determines whether [is open changed] [the specified d].
        /// </summary>
        /// <param name="d">The d.</param>
        /// <param name="e">The <see cref="DependencyPropertyChangedEventArgs"/> instance containing the event data.</param>
        private static void IsOpenChangedAction(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (IsOpenChanged != null)
            {
                IsOpenChanged.Invoke(d, e);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GridWidthOpen"/> class.
        /// </summary>
        public CustomGridPanel()
        {
            IsOpenChanged += (sender, e) =>
            {
                InvokeAction(IsOpen ? Action.Open : Action.Close);
            };
        }

        /// <summary>
        /// Raises the <see cref="E:Initialized" /> event.
        /// </summary>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);

            // Init variable
            exspectedMaxWidth = Width;
            exspectedMinWidth = MinWidth;
        }

        /// <summary>
        /// Invokes the action.
        /// </summary>
        /// <param name="action">The action.</param>
        private void InvokeAction(Action action)
        {
            if (IsLoaded)
            {
                dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
                dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 3);

                // Reset speed
                speed = DEFAULT_SPEED;

                if (action == Action.Open)
                {
                    // Open animation
                    OpenAnimation();
                }
                else
                {
                    // Close animation
                    CloseAnimation();
                }

                // Start timer
                dispatcherTimer.Start();
            }
            else if (action == Action.Close)
            {
                Visibility = Visibility.Collapsed;
                Width = exspectedMinWidth;
            }
            else if (action == Action.Open)
            {
                Visibility = Visibility.Visible;
                Width = exspectedMaxWidth;
            }
        }

        /// <summary>
        /// Closes the animation.
        /// </summary>
        private void CloseAnimation()
        {
            dispatcherTimer.Tick += (sender, e) =>
            {
                if (Width > exspectedMinWidth)
                {
                    speed += 1;

                    if (Width - exspectedMinWidth < speed)
                    {
                        speed = Width - exspectedMinWidth;
                    }

                    Width -= speed;
                }
                else
                {
                    // Stop timer
                    dispatcherTimer.Stop();
                    Visibility = Visibility.Collapsed;
                }
            };
        }

        /// <summary>
        /// Settings the panel animation.
        /// </summary>
        /// <param name="isOpenSetting">if set to <c>true</c> [is open setting].</param>
        private void OpenAnimation()
        {
            Visibility = Visibility.Visible;

            dispatcherTimer.Tick += (sender, e) =>
            {
                if (Width < exspectedMaxWidth)
                {
                    speed += 1;

                    if (exspectedMaxWidth - Width < speed)
                    {
                        speed = exspectedMaxWidth - Width;
                    }

                    Width += speed;
                }
                else
                {
                    // Stop timer
                    dispatcherTimer.Stop();
                }
            };
        }
    }
}
