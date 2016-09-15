using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Kinect;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using KinectMouseController;
using Coding4Fun.Kinect.Wpf;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls;



namespace PowerPointAddIn2
{
    public partial class RubKinect
    {
        
        Window fenetre = new Window();
        System.Windows.Controls.Image ImgKinect;
        Form fenetre1 = new Form();
        
        
        
        private void RubKinect_Load(object sender, RibbonUIEventArgs e){}
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            setupKinect();
        }

        KinectSensor nui;                                                                       //Déclaration d'un élément Capteur
        private void setupKinect()
        {
            if (KinectSensor.KinectSensors.Count == 0) { }                                        //Si non détection de la caméra
            //Alors Rien

            else
            {
                nui = KinectSensor.KinectSensors[0];                                            //Attribution du premier capteur détecté
                nui.Start();                                                                    //Démarrer le capteur kinect
                nui.SkeletonStream.Enable(new TransformSmoothParameters()
                {                                                                               //Réglage des paramètres de la caméra
                    Smoothing = 0.5f,                                                           //Pour une meilleure fluidité
                    Correction = 0.5f,
                    Prediction = 0.5f,
                    JitterRadius = 0.05f,
                    MaxDeviationRadius = 0.04f
                });
                nui.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(nui_SkeletonFrameReady);
                                                                                                //Appel de l'événement Création d'un squelette

               // nui.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);             //Définition de la résolution image
               // nui.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(ColorImageReady); 
                                                                                                //Appel de l'événement ColorImageReady

            }
        }


        bool laserOn = false;                                                                   //Booléen pour l'activation de la main en tant que curseur
        void nui_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)               //Création de l'événement
        {
            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())                         //Déclaration d'une trame pour squelette
            {                                                                                   
                if (skeletonFrame != null)
                {                                                                               //Déclaration d'un tableau de type Squelette
                                                                                                //Copie du tableau dans la trame
                                                                                                //Déclaration d'un squelette dans lequel 
                                                                                                //on introduit le premier squelette détecté
                    Skeleton[] skeletonData = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(skeletonData);
                    Skeleton playerSkeleton = (from s in skeletonData where s.TrackingState == SkeletonTrackingState.Tracked select s).FirstOrDefault();
                                                        
                    if (playerSkeleton != null)
                    {                                                                           //Appel des fonctions "Détecter un mouvement"
                        DiapoSuivante(playerSkeleton.Joints[JointType.HandRight], playerSkeleton.Joints[JointType.HipLeft]);
                        DiapoPrécédente(playerSkeleton.Joints[JointType.HandLeft], playerSkeleton.Joints[JointType.HipRight]);
                        LancerDiapo(playerSkeleton.Joints[JointType.HandRight], playerSkeleton.Joints[JointType.Head]);
                        QuitterDiapo(playerSkeleton.Joints[JointType.HandLeft], playerSkeleton.Joints[JointType.KneeLeft]);
                        ActiverLaser(playerSkeleton.Joints[JointType.HandLeft], playerSkeleton.Joints[JointType.ShoulderLeft]);               
                                                        
                                                                                                //Détection d'un mouvement servant à activer le booléen
                        if (playerSkeleton.Joints[JointType.HandLeft].Position.Y > playerSkeleton.Joints[JointType.ShoulderLeft].Position.Y)
                        {
                            if(laserOn==false)laserOn = true;
                            else if (laserOn==true)laserOn = false;
                        } 
                                                                                                //Si booléen activé alors détection de la main droite en tant que curseur
                        if(laserOn==true)
                        {

                            Joint scaledRight = playerSkeleton.Joints[JointType.HandRight].ScaleTo((int)SystemInformation.PrimaryMonitorSize.Width, (int)SystemInformation.PrimaryMonitorSize.Height, -playerSkeleton.Position.X, -playerSkeleton.Position.Y);
                            int CursorX = (int)scaledRight.Position.X;
                            int CursorY = (int)scaledRight.Position.Y;
                            KinectMouseController.KinectMouseMethods.SendMouseInput(CursorX, CursorY, SystemInformation.PrimaryMonitorSize.Width, SystemInformation.PrimaryMonitorSize.Height, false);
                            
                        }   
                    }
                }
            }
        }

        int comptSuiv;                                              //Entier servant de tempo entre deux gestes
        private void DiapoSuivante(Joint handRight, Joint hipLeft)
        {
            if (handRight.Position.X < hipLeft.Position.X)          //Si main droite à gauche de hanche gauche
            {

                comptSuiv++;                                        //Alors incrémentation du compteur
            }

            if (comptSuiv == 20)                                    //Si compteur arrivé à 20
            {
                System.Windows.Forms.SendKeys.SendWait("{Right}");
                comptSuiv = 0;                                      //Alors envoie de la touche "Droite"
            }                                                       //Et remise à 0 du compteur
        }

        int comptPrec;                                              //Entier servant de tempo entre deux gestes
        private void DiapoPrécédente(Joint handLeft, Joint hipRight)
        {
            if (handLeft.Position.X > hipRight.Position.X)          //Si main gauche à droite de hanche droite
            {

                comptPrec++;                                        //Alors incrémentation du compteur
            }

            if (comptPrec == 20)                                    //Si compteur arrivé à 20
            {
                System.Windows.Forms.SendKeys.SendWait("{Left}");   //Alors envoie de la touche "Gauche"
                comptPrec = 0;                                      //Et remise à 0 du compteur
            }
        }

        int comptLanc;
        private void LancerDiapo(Joint handRight, Joint head)
        {
            if(handRight.Position.Y > head.Position.Y)              //Si main droite en haut de tête
            {
                comptLanc++;
            }

            if(comptLanc==50)
            {
                System.Windows.Forms.SendKeys.SendWait("{F5}");     //Envoie de touche F5
                comptLanc = 0;
            }
        }

        int comptQuit;
        private void QuitterDiapo(Joint handLeft, Joint kneeLeft)
        {
            if(handLeft.Position.Y < kneeLeft.Position.Y)           //Si main gauche en bas du genou gauche
            {
                comptQuit++;
            }

            if(comptQuit==20)
            {
                System.Windows.Forms.SendKeys.SendWait("{ESC}");    //Evoie de touche echap
                comptQuit = 0;
            }
        }

        int comptLaser;
        private void ActiverLaser(Joint handLeft, Joint shoulderLeft)
        {
            if(handLeft.Position.Y > shoulderLeft.Position.Y)       //Si main gauche en haut de épaule gauche
            {
                comptLaser++;
            }

            if(comptLaser==20)
            {
                System.Windows.Forms.SendKeys.SendWait("^l");       //Envoie de "CTRL+L"
                comptLaser = 0;
            }
        }

        byte[] pixelData;



        void ColorImageReady(object sender, ColorImageFrameReadyEventArgs e)    //Utilisation d'un Argument de l'évènement "Cadre prêt"
        {
            bool receivedData = false; //booléen utilisé pour vérifié que l'on a bien reçu des données
            using (ColorImageFrame colorImageFrame = e.OpenColorImageFrame())   //Déclaration d'un cadre d'image auquel on affecte 
            {
                if (colorImageFrame != null) //Si le cadre reçoit des informations alors 
                {
                    if (pixelData == null)  //On initialise notre octet pixelData
                    {
                        pixelData = new byte[colorImageFrame.PixelDataLength];  //On lui affecte la longueur des pixels de notre cadre récupéré 
                    }
                    colorImageFrame.CopyPixelDataTo(pixelData); //On copie les données des pixels du cadre dans les notre Octet
                    receivedData = true;
                }
                else
                {

                }

                if (receivedData) //Vérification que données reçues
                {
                    ImageSource imgSrc = BitmapSource.Create(colorImageFrame.Width, colorImageFrame.Height, 96, 96, PixelFormats.Bgr32, null, pixelData, colorImageFrame.Width * colorImageFrame.BytesPerPixel);
                    ImgKinect.Source = imgSrc;

                }
            }
        }

      
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            nui.Stop();                                             //Arrêt du capteur    
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            /*setup2();
            fenetre1.Show();
            PictureBox pictureBox2 = new PictureBox();
            pictureBox2.Image = ImgKinect;                              //Pas bonne bilbliothèque
            pictureBox2.Show();
            fenetre1.Controls.Add(pictureBox2);
            pictureBox2.Dock = System.Windows.Forms.DockStyle.Fill;*/  //Appel de la fenêtre où est contenu l'image
                                                                       //Pas encore au point
        }

        
    }
}
