// Chorus Crisp - Vegas Pro 22 Script
// For Sparta Remix vocal chop layering effect

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    const double MIN_SPLICE_TIME = 0.020;
    const double MAX_SPLICE_TIME = 0.047;
    const double MIN_DUCK_DB = -1.0;
    const double MAX_DUCK_DB = -17.0;

    public void FromVegas(Vegas vegas)
    {
        List<TrackEvent> selectedEvents = new List<TrackEvent>();
        
        foreach (Track track in vegas.Project.Tracks)
        {
            foreach (TrackEvent ev in track.Events)
            {
                if (ev.Selected && ev.IsAudio())
                {
                    selectedEvents.Add(ev);
                }
            }
        }

        if (selectedEvents.Count == 0)
        {
            MessageBox.Show("No audio clips selected!\nSelect some audio events and try again.", 
                "Chorus Crisp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ChorusCrispDialog dialog = new ChorusCrispDialog();
        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        double splicePercent = dialog.SplicePosition / 100.0;
        double crispPercent = dialog.Crispness / 100.0;
        CurveType fadeType = dialog.SelectedCurveType;

        double spliceTime = MIN_SPLICE_TIME + (splicePercent * (MAX_SPLICE_TIME - MIN_SPLICE_TIME));
        double duckDb = MIN_DUCK_DB + (crispPercent * (MAX_DUCK_DB - MIN_DUCK_DB));

        int successCount = 0;
        int errorCount = 0;
        string lastError = "";

        using (UndoBlock undo = new UndoBlock("Chorus Crisp"))
        {
            foreach (TrackEvent ev in selectedEvents)
            {
                try
                {
                    ProcessEvent(ev, spliceTime, duckDb, fadeType);
                    successCount++;
                }
                catch (Exception ex)
                {
                    errorCount++;
                    lastError = ex.Message;
                }
            }
        }

        string message = String.Format("Processed {0} clip(s)!\n\nSplice at: {1:F3}s\nVolume duck: {2:F1} dB\nFade type: {3}",
            successCount, spliceTime, duckDb, fadeType);
        if (errorCount > 0)
        {
            message += String.Format("\n\n{0} clip(s) had errors:\n{1}", errorCount, lastError);
        }
        MessageBox.Show(message, "Chorus Crisp Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ProcessEvent(TrackEvent originalEvent, double spliceTime, 
        double duckDb, CurveType fadeType)
    {
        Timecode eventLength = originalEvent.Length;
        Timecode spliceOffset = Timecode.FromSeconds(spliceTime);

        if (eventLength.ToMilliseconds() < spliceTime * 1000 + 10)
        {
            return;
        }

        TrackEvent secondEvent = originalEvent.Split(spliceOffset);

        if (secondEvent == null)
            return;

        Timecode overlapDuration = spliceOffset;

        AudioEvent audioSecond = secondEvent as AudioEvent;
        if (audioSecond != null)
        {
            Timecode newStart = secondEvent.Start - overlapDuration;
            
            foreach (Take take in audioSecond.Takes)
            {
                take.Offset = take.Offset - overlapDuration;
            }
            
            secondEvent.Start = newStart;
            secondEvent.Length = secondEvent.Length + overlapDuration;
            
            double linearGain = Math.Pow(10.0, duckDb / 20.0);
            audioSecond.NormalizeGain = linearGain;
        }

        secondEvent.FadeIn.Length = overlapDuration;
        secondEvent.FadeIn.Curve = fadeType;

        originalEvent.FadeOut.Length = overlapDuration;
        originalEvent.FadeOut.Curve = fadeType;
    }
}

public class ChorusCrispDialog : Form
{
    private TrackBar spliceSlider;
    private TrackBar crispSlider;
    private ComboBox curveCombo;
    private Label spliceLabel;
    private Label crispLabel;
    private Button okButton;
    private Button cancelButton;

    public int SplicePosition { get { return spliceSlider.Value; } }
    public int Crispness { get { return crispSlider.Value; } }
    public CurveType SelectedCurveType 
    { 
        get 
        {
            switch (curveCombo.SelectedIndex)
            {
                case 0: return CurveType.Linear;
                case 1: return CurveType.Fast;
                case 2: return CurveType.Slow;
                case 3: return CurveType.Sharp;
                case 4: return CurveType.Smooth;
                default: return CurveType.Linear;
            }
        }
    }

    public ChorusCrispDialog()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.AutoScaleDimensions = new SizeF(96F, 96F);
        
        this.Text = "Chorus Crisp";
        this.ClientSize = new Size(400, 330);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.FromArgb(45, 45, 48);
        this.ForeColor = Color.White;

        int yPos = 20;

        Label spliceTitle = new Label();
        spliceTitle.Text = "Splice Position (0.020s - 0.047s):";
        spliceTitle.Location = new Point(20, yPos);
        spliceTitle.Size = new Size(280, 20);
        spliceTitle.ForeColor = Color.White;
        this.Controls.Add(spliceTitle);

        yPos += 28;

        spliceSlider = new TrackBar();
        spliceSlider.Location = new Point(20, yPos);
        spliceSlider.Size = new Size(280, 45);
        spliceSlider.Minimum = 0;
        spliceSlider.Maximum = 100;
        spliceSlider.Value = 52;
        spliceSlider.TickFrequency = 10;
        spliceSlider.ValueChanged += SpliceSlider_ValueChanged;
        this.Controls.Add(spliceSlider);

        spliceLabel = new Label();
        spliceLabel.Location = new Point(310, yPos + 8);
        spliceLabel.Size = new Size(70, 20);
        spliceLabel.ForeColor = Color.FromArgb(100, 200, 255);
        spliceLabel.Text = "0.034s";
        this.Controls.Add(spliceLabel);

        yPos += 55;

        Label crispTitle = new Label();
        crispTitle.Text = "Crispness / Volume Duck (-1 to -17 dB):";
        crispTitle.Location = new Point(20, yPos);
        crispTitle.Size = new Size(280, 20);
        crispTitle.ForeColor = Color.White;
        this.Controls.Add(crispTitle);

        yPos += 28;

        crispSlider = new TrackBar();
        crispSlider.Location = new Point(20, yPos);
        crispSlider.Size = new Size(280, 45);
        crispSlider.Minimum = 0;
        crispSlider.Maximum = 100;
        crispSlider.Value = 0;
        crispSlider.TickFrequency = 10;
        crispSlider.ValueChanged += CrispSlider_ValueChanged;
        this.Controls.Add(crispSlider);

        crispLabel = new Label();
        crispLabel.Location = new Point(310, yPos + 8);
        crispLabel.Size = new Size(70, 20);
        crispLabel.ForeColor = Color.FromArgb(100, 200, 255);
        crispLabel.Text = "-1.0 dB";
        this.Controls.Add(crispLabel);

        yPos += 55;

        Label curveTitle = new Label();
        curveTitle.Text = "Crossfade Curve:";
        curveTitle.Location = new Point(20, yPos);
        curveTitle.Size = new Size(130, 20);
        curveTitle.ForeColor = Color.White;
        this.Controls.Add(curveTitle);

        yPos += 25;

        curveCombo = new ComboBox();
        curveCombo.Location = new Point(20, yPos);
        curveCombo.Size = new Size(360, 28);
        curveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        curveCombo.Font = new Font(curveCombo.Font.FontFamily, 10);
        curveCombo.Items.AddRange(new string[] { "Linear", "Fast", "Slow", "Sharp", "Smooth" });
        curveCombo.SelectedIndex = 0;
        this.Controls.Add(curveCombo);

        yPos += 45;

        okButton = new Button();
        okButton.Text = "Apply Crisp";
        okButton.Location = new Point(70, yPos);
        okButton.Size = new Size(120, 35);
        okButton.BackColor = Color.FromArgb(0, 122, 204);
        okButton.ForeColor = Color.White;
        okButton.FlatStyle = FlatStyle.Flat;
        okButton.DialogResult = DialogResult.OK;
        this.Controls.Add(okButton);

        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(210, yPos);
        cancelButton.Size = new Size(120, 35);
        cancelButton.BackColor = Color.FromArgb(80, 80, 80);
        cancelButton.ForeColor = Color.White;
        cancelButton.FlatStyle = FlatStyle.Flat;
        cancelButton.DialogResult = DialogResult.Cancel;
        this.Controls.Add(cancelButton);

        this.AcceptButton = okButton;
        this.CancelButton = cancelButton;
    }

    private void SpliceSlider_ValueChanged(object sender, EventArgs e)
    {
        double spliceTime = 0.020 + (spliceSlider.Value / 100.0) * (0.047 - 0.020);
        spliceLabel.Text = String.Format("{0:F3}s", spliceTime);
    }

    private void CrispSlider_ValueChanged(object sender, EventArgs e)
    {
        double duckDb = -1.0 + (crispSlider.Value / 100.0) * (-17.0 - (-1.0));
        crispLabel.Text = String.Format("{0:F1} dB", duckDb);
    }
}
