// Standard .NET Framework namespaces for WPF (Windows Presentation Foundation) UI components
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
// Microsoft Kinect SDK namespace - provides access to Kinect sensor hardware
using Microsoft.Kinect;
// Microsoft Speech Recognition namespace - for voice command recognition (if implemented)
using Microsoft.Speech.Recognition;
using System.Threading;
using System.IO;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
// DispatcherTimer is used for UI thread operations in WPF
using System.Windows.Threading;

namespace KinectPowerPointControl
{
    /// <summary>
    /// MainWindow class - The primary window of the application that handles:
    /// 1. Kinect sensor initialization and data capture
    /// 2. Skeleton tracking and gesture recognition
    /// 3. Visual feedback through colored ellipses
    /// 4. Sending keyboard commands to control PowerPoint presentations
    /// </summary>
    public partial class MainWindow : Window
    {
        // KinectSensor object - represents the physical Kinect device connected to the computer
        // This is the main interface for accessing all Kinect capabilities (color, depth, skeleton, audio)
        KinectSensor sensor;
        
        // SpeechRecognitionEngine - for processing voice commands (initialized but not fully implemented in this version)
        SpeechRecognitionEngine speechRecognizer;

        // DispatcherTimer - can be used for timed operations on the UI thread
        // Currently declared but not actively used in the current implementation
        DispatcherTimer readyTimer;

        // Byte array to store color image data from the Kinect RGB camera
        // This buffer holds the pixel data for each frame captured
        byte[] colorBytes;
        
        // Array to store skeleton tracking data - contains information about all tracked persons
        // Each Skeleton object contains joint positions and tracking states
        Skeleton[] skeletons;

        // Boolean flag to track whether the tracking circles (ellipses) are currently visible on screen
        // Users can toggle visibility with the 'C' key
        bool isCirclesVisible = true;

        // Gesture state flags - prevent multiple triggers of the same gesture
        // These ensure gestures only activate once when the threshold is crossed
        bool isForwardGestureActive = false;  // True when right hand is extended (forward/next slide)
        bool isBackGestureActive = false;     // True when left hand is extended (back/previous slide)
        
        // Visual feedback brushes - colors change to indicate gesture activation
        // Green = gesture is active, Red = gesture is inactive
        SolidColorBrush activeBrush = new SolidColorBrush(Colors.Green);
        SolidColorBrush inactiveBrush = new SolidColorBrush(Colors.Red);

        /// <summary>
        /// Constructor for MainWindow - initializes the window and sets up event handlers
        /// This is called when the application starts and the window is created
        /// </summary>
        public MainWindow()
        {
            // InitializeComponent() loads the XAML markup and creates all UI elements defined in MainWindow.xaml
            InitializeComponent();

            // Runtime initialization is handled when the window is opened. When the window
            // is closed, the runtime MUST be uninitialized to properly release Kinect resources.
            // The Loaded event fires after the window is fully loaded and ready to use
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
            // Handle the content obtained from the video camera, once received.

            // Register keyboard event handler - allows the application to respond to key presses
            // Currently used to toggle circle visibility with the 'C' key
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        /// <summary>
        /// Event handler called when the MainWindow is fully loaded and ready
        /// This method initializes the Kinect sensor and sets up all data streams
        /// </summary>
        /// <param name="sender">The object that raised the event (MainWindow)</param>
        /// <param name="e">Event arguments containing information about the Loaded event</param>
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Get the first available Kinect sensor from the collection of connected sensors
            // FirstOrDefault() returns the first sensor or null if none are connected
            sensor = KinectSensor.KinectSensors.FirstOrDefault();

            // Check if a Kinect sensor was found - if not, show error and close application
            if (sensor == null)
            {
                MessageBox.Show("This application requires a Kinect sensor.");
                this.Close();
                return; // Exit early if no sensor found
            }

            // Start the Kinect sensor - this activates the hardware and begins data collection
            sensor.Start();

            // Enable the color stream (RGB camera) at 640x480 resolution, 30 frames per second
            // This provides the video feed that users see in the application window
            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            // Register event handler - called each time a new color frame is available
            sensor.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(sensor_ColorFrameReady);

            // Enable the depth stream at 320x240 resolution, 30 frames per second
            // Depth data is used internally by the skeleton tracking system to determine 3D positions
            // Note: We don't directly process depth frames, but skeleton tracking requires depth data
            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);

            // Enable skeleton tracking - this allows the Kinect to detect and track human bodies
            // The skeleton stream provides 3D positions of 20 body joints (head, shoulders, hands, etc.)
            sensor.SkeletonStream.Enable();
            // Register event handler - called each time skeleton tracking data is available
            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            // Optional: Adjust the Kinect's tilt angle (elevation) to point up or down
            // Uncomment and adjust the value (-27 to +27 degrees) if needed for better tracking
            //sensor.ElevationAngle = 10;

            // Register exit event handler - ensures proper cleanup when application closes
            // This is critical to prevent resource leaks and properly shut down the Kinect
            Application.Current.Exit += new ExitEventHandler(Current_Exit);

            // Initialize speech recognition (if implemented)
            // Note: This method is called but the implementation is not present in this codebase.
            // If you want to add speech recognition, you would need to implement this method
            // to set up the SpeechRecognitionEngine and configure voice commands.
            InitializeSpeechRecognition();
        }

        /// <summary>
        /// Event handler called when the application is exiting
        /// Performs cleanup operations to properly release Kinect resources
        /// This is essential to prevent resource leaks and ensure the sensor can be used by other applications
        /// </summary>
        /// <param name="sender">The Application object that raised the event</param>
        /// <param name="e">Event arguments containing information about the Exit event</param>
        void Current_Exit(object sender, ExitEventArgs e)
        {
            // Clean up speech recognition resources if they were initialized
            if (speechRecognizer != null)
            {
                // Cancel any ongoing speech recognition operations
                speechRecognizer.RecognizeAsyncCancel();
                // Stop the asynchronous recognition engine
                speechRecognizer.RecognizeAsyncStop();
            }
            
            // Clean up Kinect sensor resources
            if (sensor != null)
            {
                // Stop the audio source if it was started
                sensor.AudioSource.Stop();
                // Stop the sensor - this releases hardware resources
                sensor.Stop();
                // Dispose of the sensor object to free up memory
                sensor.Dispose();
                // Set to null to indicate cleanup is complete
                sensor = null;
            }
        }

        /// <summary>
        /// Event handler for keyboard key press events
        /// Allows users to interact with the application using keyboard shortcuts
        /// </summary>
        /// <param name="sender">The object that raised the event</param>
        /// <param name="e">KeyEventArgs containing information about which key was pressed</param>
        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            // Check if the 'C' key was pressed
            // This toggles the visibility of the tracking circles (head and hands)
            if (e.Key == Key.C)
            {
                ToggleCircles();
            }
        }

        /// <summary>
        /// Event handler called when a new color (RGB) frame is available from the Kinect camera
        /// This method processes the color image data and displays it in the application window
        /// </summary>
        /// <param name="sender">The KinectSensor object that raised the event</param>
        /// <param name="e">ColorImageFrameReadyEventArgs containing the new color frame data</param>
        void sensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            // Open the color image frame - using statement ensures proper disposal
            // The frame must be disposed quickly to avoid memory issues
            using (var image = e.OpenColorImageFrame())
            {
                // Check if frame is valid - sometimes frames can be null if sensor is busy
                if (image == null)
                    return;

                // Allocate or resize the colorBytes buffer if needed
                // This ensures we have enough space to store the pixel data
                if (colorBytes == null ||
                    colorBytes.Length != image.PixelDataLength)
                {
                    colorBytes = new byte[image.PixelDataLength];
                }

                // Copy the raw pixel data from the Kinect frame into our buffer
                // This gives us access to the RGB color values for each pixel
                image.CopyPixelDataTo(colorBytes);

                // You could use PixelFormats.Bgr32 below to ignore the alpha channel,
                // or if you need to set the alpha you would loop through the bytes 
                // as in this loop below
                // The color data comes in BGRA format (Blue, Green, Red, Alpha)
                // We set the alpha channel (transparency) to 255 (fully opaque) for all pixels
                int length = colorBytes.Length;
                for (int i = 0; i < length; i += 4)
                {
                    // Each pixel is 4 bytes: B, G, R, A
                    // Index i+3 is the alpha channel (4th byte of each pixel)
                    colorBytes[i + 3] = 255; // Set alpha to fully opaque
                }

                // Create a BitmapSource from the color data - this is what WPF uses to display images
                // Parameters:
                //   - image.Width, image.Height: Dimensions of the image (640x480)
                //   - 96, 96: DPI (dots per inch) for the image
                //   - PixelFormats.Bgra32: Format indicating 32 bits per pixel (8 bits per channel)
                //   - null: Palette (not needed for RGB images)
                //   - colorBytes: The actual pixel data
                //   - image.Width * image.BytesPerPixel: Stride (bytes per row)
                BitmapSource source = BitmapSource.Create(image.Width,
                    image.Height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    colorBytes,
                    image.Width * image.BytesPerPixel);
                
                // Display the image in the videoImage control defined in MainWindow.xaml
                videoImage.Source = source;
            }
        }
        
        /// <summary>
        /// Event handler called when new skeleton tracking data is available from the Kinect
        /// This is the core method that processes body tracking and gesture recognition
        /// It identifies the closest person, extracts joint positions, and processes gestures
        /// </summary>
        /// <param name="sender">The KinectSensor object that raised the event</param>
        /// <param name="e">SkeletonFrameReadyEventArgs containing skeleton tracking data</param>
        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            // Open the skeleton frame - using statement ensures proper disposal
            using (var skeletonFrame = e.OpenSkeletonFrame())
            {
                // Check if frame is valid
                if (skeletonFrame == null)
                    return;

                // Allocate or resize the skeletons array if needed
                // The array size depends on how many people the Kinect can track simultaneously (up to 6)
                if (skeletons == null ||
                    skeletons.Length != skeletonFrame.SkeletonArrayLength)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                }

                // Copy skeleton data from the frame into our array
                // This gives us access to all tracked skeletons in the current frame
                skeletonFrame.CopySkeletonDataTo(skeletons);
            }

            // Find the closest fully tracked skeleton (person) to the Kinect
            // This ensures we control PowerPoint based on the person closest to the sensor
            // The query:
            //   1. Filters to only skeletons that are fully tracked (not just position-only)
            //   2. Orders by distance (Z) and horizontal position (X) - closest and most centered
            //   3. Takes the first result (closest person)
            Skeleton closestSkeleton = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                .FirstOrDefault();

            // If no tracked skeleton found, exit early (no one is being tracked)
            if (closestSkeleton == null)
                return;

            // Extract the three key joints we need for gesture recognition:
            // Head: Used as a reference point to measure hand extension
            // Right Hand: Controls forward/next slide gesture
            // Left Hand: Controls back/previous slide gesture
            var head = closestSkeleton.Joints[JointType.Head];
            var rightHand = closestSkeleton.Joints[JointType.HandRight];
            var leftHand = closestSkeleton.Joints[JointType.HandLeft];

            // Verify that all three joints are being tracked properly
            // If any joint is not tracked, we cannot reliably detect gestures
            if (head.TrackingState == JointTrackingState.NotTracked ||
                rightHand.TrackingState == JointTrackingState.NotTracked ||
                leftHand.TrackingState == JointTrackingState.NotTracked)
            {
                // Don't have a good read on the joints so we cannot process gestures
                return;
            }

            // Update the visual indicators (ellipses) on screen to show joint positions
            // Parameters: ellipse control, joint data, whether gesture is active (affects color/size)
            SetEllipsePosition(ellipseHead, head, false);  // Head is never highlighted
            SetEllipsePosition(ellipseLeftHand, leftHand, isBackGestureActive);   // Green when back gesture active
            SetEllipsePosition(ellipseRightHand, rightHand, isForwardGestureActive); // Green when forward gesture active

            // Process gestures to detect if user wants to go forward or back in PowerPoint
            // This method checks hand positions relative to head and sends keyboard commands
            ProcessForwardBackGesture(head, rightHand, leftHand);
        }

        /// <summary>
        /// This method is used to position the ellipses on the canvas according to correct movements of the tracked joints.
        /// It converts 3D skeleton coordinates to 2D screen coordinates and updates the visual indicators.
        /// The ellipses grow and change color when gestures are active to provide visual feedback.
        /// </summary>
        /// <param name="ellipse">The WPF Ellipse control to position (head, left hand, or right hand)</param>
        /// <param name="joint">The Kinect joint data containing 3D position information</param>
        /// <param name="isHighlighted">True if the gesture is currently active (makes ellipse larger and green)</param>
        private void SetEllipsePosition(Ellipse ellipse, Joint joint, bool isHighlighted)
        {
            // Update ellipse size and color based on gesture state
            if (isHighlighted)
            {
                // When gesture is active: make ellipse larger (60x60) and green
                // This provides clear visual feedback that a gesture has been detected
                ellipse.Width = 60;
                ellipse.Height = 60;
                ellipse.Fill = activeBrush; // Green color
            }
            else
            {
                // When gesture is inactive: make ellipse smaller (20x20) and red
                // Normal tracking state
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = inactiveBrush; // Red color
            }

            // Get the CoordinateMapper - this converts between different coordinate systems
            // Kinect uses 3D skeleton space (meters), but we need 2D color image space (pixels)
            CoordinateMapper mapper = sensor.CoordinateMapper;

            // Convert the 3D skeleton joint position to a 2D point in the color image
            // This maps the joint's position in 3D space to where it appears in the RGB camera view
            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            // Position the ellipse on the canvas, centered on the joint position
            // We subtract half the ellipse width/height to center it on the joint point
            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }
        
        /// <summary>
        /// Processes forward and back gestures by comparing hand positions to head position.
        /// This is the core gesture recognition logic that detects when a user extends their arm.
        /// 
        /// Gesture Detection Logic:
        /// - Forward Gesture: Right hand extends 45cm (0.45 meters) to the right of the head
        ///   → Sends Right Arrow key → Advances PowerPoint to next slide
        /// - Back Gesture: Left hand extends 45cm (0.45 meters) to the left of the head
        ///   → Sends Left Arrow key → Goes back to previous slide in PowerPoint
        /// 
        /// The gesture only triggers once when crossing the threshold, preventing multiple rapid triggers.
        /// The user must bring their hand back closer to reset the gesture before it can trigger again.
        /// </summary>
        /// <param name="head">The head joint - used as reference point for measuring hand extension</param>
        /// <param name="rightHand">The right hand joint - controls forward/next slide gesture</param>
        /// <param name="leftHand">The left hand joint - controls back/previous slide gesture</param>
        private void ProcessForwardBackGesture(Joint head, Joint rightHand, Joint leftHand)
        {
            // FORWARD GESTURE DETECTION (Right Hand)
            // Check if right hand is extended more than 45cm (0.45 meters) to the right of the head
            // Position values are in meters in 3D space relative to the Kinect sensor
            if (rightHand.Position.X > head.Position.X + 0.45)
            {
                // Only trigger if the gesture wasn't already active (prevents multiple triggers)
                if (!isForwardGestureActive)
                {
                    // Mark gesture as active
                    isForwardGestureActive = true;
                    // Send Right Arrow key to the foreground application (PowerPoint)
                    // This simulates pressing the Right Arrow key on the keyboard
                    // PowerPoint interprets this as "next slide" command
                    System.Windows.Forms.SendKeys.SendWait("{Right}");
                }
            }
            else
            {
                // Hand is back within normal range - reset gesture state
                // This allows the gesture to be triggered again when hand extends again
                isForwardGestureActive = false;
            }

            // BACK GESTURE DETECTION (Left Hand)
            // Check if left hand is extended more than 45cm (0.45 meters) to the left of the head
            // Negative X values mean to the left of the sensor's center
            if (leftHand.Position.X < head.Position.X - 0.45)
            {
                // Only trigger if the gesture wasn't already active
                if (!isBackGestureActive)
                {
                    // Mark gesture as active
                    isBackGestureActive = true;
                    // Send Left Arrow key to the foreground application (PowerPoint)
                    // This simulates pressing the Left Arrow key on the keyboard
                    // PowerPoint interprets this as "previous slide" command
                    System.Windows.Forms.SendKeys.SendWait("{Left}");
                }
            }
            else
            {
                // Hand is back within normal range - reset gesture state
                isBackGestureActive = false;
            }
        }
        
        /// <summary>
        /// Toggles the visibility of the tracking circles (ellipses) on and off.
        /// Called when the user presses the 'C' key.
        /// This allows users to hide the visual indicators if they find them distracting.
        /// </summary>
        void ToggleCircles()
        {
            // If circles are currently visible, hide them; otherwise show them
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }

        /// <summary>
        /// Hides all tracking circles (head, left hand, right hand) from the display.
        /// The circles are set to Collapsed, which removes them from the visual tree entirely.
        /// </summary>
        void HideCircles()
        {
            // Update the visibility flag
            isCirclesVisible = false;
            // Hide all three tracking ellipses
            ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightHand.Visibility = System.Windows.Visibility.Collapsed;
        }

        /// <summary>
        /// Shows all tracking circles (head, left hand, right hand) on the display.
        /// Makes the visual indicators visible so users can see where the Kinect is tracking.
        /// </summary>
        void ShowCircles()
        {
            // Update the visibility flag
            isCirclesVisible = true;
            // Show all three tracking ellipses
            ellipseHead.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Visible;
            ellipseRightHand.Visibility = System.Windows.Visibility.Visible;
        }

        /// <summary>
        /// Makes the window visible and brings it to the front.
        /// Sets the window to be topmost (always on top) and maximizes it to full screen.
        /// This can be useful for presentations where you want the application to be clearly visible.
        /// </summary>
        private void ShowWindow()
        {
            // Make window stay on top of all other windows
            this.Topmost = true;
            // Maximize the window to fill the entire screen
            this.WindowState = System.Windows.WindowState.Maximized;
        }

        /// <summary>
        /// Hides the window by minimizing it and removing the topmost property.
        /// This allows users to minimize the application window if needed.
        /// </summary>
        private void HideWindow()
        {
            // Remove topmost property so window can be behind others
            this.Topmost = false;
            // Minimize the window to the taskbar
            this.WindowState = System.Windows.WindowState.Minimized;
        }
        
        /// <summary>
        /// Placeholder method for initializing speech recognition functionality.
        /// NOTE: This method is currently not implemented. The method call exists in MainWindow_Loaded,
        /// but the actual implementation is missing. To add speech recognition:
        /// 1. Configure the SpeechRecognitionEngine with grammar/commands
        /// 2. Set up event handlers for speech recognition events
        /// 3. Start the recognition engine using RecognizeAsync()
        /// </summary>
        private void InitializeSpeechRecognition()
        {
            // Speech recognition initialization code would go here
            // This is a placeholder to prevent compilation errors
            // The speechRecognizer field is declared but not used in the current implementation
        }

    }
}
