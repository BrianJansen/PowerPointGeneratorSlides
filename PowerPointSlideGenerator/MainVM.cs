using Google.Apis.Services;
using Google.Apis.Customsearch.v1.Data;
using Google.Apis.Customsearch.v1;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Documents;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointSlideGenerator
{
    public class MainVM : BindableBase
    {
        int index = 0;
        List<string> boldStrings;
        // APIKey: 
        // AIzaSyCug7MHldoDpThh1Q6OqOqkFSm7yLf5ZcQ
        // AIzaSyAoZCKZ4XtrXyXhqvyf8-STyiZqi6Z0BPM
        // AIzaSyDoZrnoUkLitks-1dNshwnWb39EDHsvcX8
        // 

        Dictionary<string, string> thumbnailFullsizeDict = new Dictionary<string, string>();

        ObservableCollection<BitmapImage> _Images = new ObservableCollection<BitmapImage>();
        public ObservableCollection<BitmapImage> Images { get => _Images; set => SetProperty(ref _Images, value); }

        ObservableCollection<BitmapImage> _ConfirmedImages = new ObservableCollection<BitmapImage>();
        public ObservableCollection<BitmapImage> ConfirmedImages { get => _ConfirmedImages; set => SetProperty(ref _ConfirmedImages, value); }

        public MainVM()
        {
            SearchImagesCmd = new DelegateCommand(OnSearchImagesCmd);
            BoldSelectedCmd = new DelegateCommand(OnBoldSelectedCmd);
            GenerateSlideCmd = new DelegateCommand(OnGenerateSlideCmd);
            ImageSelected = new DelegateCommand<BitmapImage>(OnImageSelected);
            ConfirmedImageSelected = new DelegateCommand<BitmapImage>(OnConfirmedImageSelected);

            boldStrings = new List<string>();
        }
        
        public RichTextBox TextArea { get; set; }
        string _TitleAreaText;
        public string TitleAreaText
        {
            get { return _TitleAreaText; }
            set { index = 0; SetProperty(ref _TitleAreaText, value); }
        }
        public DelegateCommand GenerateSlideCmd { get; set; }
        public DelegateCommand BoldSelectedCmd { get; set; }
        public DelegateCommand SearchImagesCmd { get; set; }
        public DelegateCommand<BitmapImage> ImageSelected { get; set; }
        public DelegateCommand<BitmapImage> ConfirmedImageSelected { get; set; }

        void OnBoldSelectedCmd()
        {
            if (TextArea.Selection.GetPropertyValue(TextElement.FontWeightProperty).Equals(FontWeights.Bold))
            {
                boldStrings.Remove(TextArea.Selection.Text);
                TextArea.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
            }
            else
            {
                boldStrings.Add(TextArea.Selection.Text);
                TextArea.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
            }
        }

        void OnSearchImagesCmd()
        {
            CustomsearchService svc = new CustomsearchService(new BaseClientService.Initializer
            {
                ApplicationName = "PowerPointSlideGenerator",
                ApiKey = "AIzaSyDoZrnoUkLitks-1dNshwnWb39EDHsvcX8",
            });


            string searchStr = TitleAreaText + " " + String.Join(" ", boldStrings);

            Google.Apis.Customsearch.v1.CseResource.ListRequest listRequest = svc.Cse.List();
            listRequest.Start = index;
            listRequest.Q = searchStr;
            listRequest.Cx = "4783681fcb1c63ffc";
            listRequest.SearchType = CseResource.ListRequest.SearchTypeEnum.Image;
            Search search = listRequest.Execute();

            index += 10;

            _Images.Clear();
            foreach (Result result1 in search.Items)
            {
                if(!thumbnailFullsizeDict.ContainsKey(result1.Image.ThumbnailLink))
                    thumbnailFullsizeDict.Add(result1.Image.ThumbnailLink, result1.Link);
                Images.Add(new BitmapImage(new Uri(result1.Image.ThumbnailLink)));
            }

        }
        void OnGenerateSlideCmd()
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptpresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
            System.Windows.Documents.TextRange bodyText = new System.Windows.Documents.TextRange(TextArea.Document.ContentStart, TextArea.Document.ContentEnd);
            for (int i = 0; i < ConfirmedImages.Count; i++)
            {
                Microsoft.Office.Interop.PowerPoint.Slides slides;
                Microsoft.Office.Interop.PowerPoint._Slide slide;
                Microsoft.Office.Interop.PowerPoint.TextRange TitleText;
                Microsoft.Office.Interop.PowerPoint.TextRange BodyText;

                Microsoft.Office.Interop.PowerPoint.CustomLayout custLayout = pptpresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

                slides = pptpresentation.Slides;
                slide = slides.AddSlide(i+1,custLayout);

                TitleText = slide.Shapes[1].TextFrame.TextRange;
                TitleText.Text = TitleAreaText;
                TitleText.Font.Name = "Arial";
                TitleText.Font.Size = 32;

                Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
                if(thumbnailFullsizeDict.ContainsKey(ConfirmedImages[i].UriSource.OriginalString))
                    slide.Shapes.AddPicture(thumbnailFullsizeDict[ConfirmedImages[i].UriSource.OriginalString], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);

                slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 70, 100, 500, 50);
                BodyText = slide.Shapes[3].TextFrame.TextRange;
                BodyText.Text = bodyText.Text;
                BodyText.Font.Name = "Arial";
                BodyText.Font.Size = 32;
            }
            try
            {
                pptpresentation.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\newslide.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch(Exception e)
            {
                //don't know how error thing need to be handled, but a likely cause of an error here would be if the file is already open and in use.
            }
        }

        void OnImageSelected( BitmapImage selectedImage)
        {
            if(ConfirmedImages.Count < 3)
            {
                ConfirmedImages.Add(selectedImage);
            }
        }

        void OnConfirmedImageSelected(BitmapImage selectedImage)
        {
            ConfirmedImages.Remove(selectedImage);
        }
    }




}
