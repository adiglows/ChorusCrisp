// Chorus Crisp - Vegas Pro Script
// For Sparta Remix vocal chop layering effect

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    const double MIN_SPLICE_TIME = 0.020;
    const double MAX_SPLICE_TIME = 0.060;
    const double MIN_DUCK_DB = 0.0;
    const double MAX_DUCK_DB = -15.0;

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
        double offsetPercent = dialog.OffsetAmount / 100.0;
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
                    ProcessEvent(ev, spliceTime, duckDb, offsetPercent, fadeType);
                    successCount++;
                }
                catch (Exception ex)
                {
                    errorCount++;
                    lastError = ex.Message;
                }
            }
        }

        string message = String.Format("Processed {0} clip(s)!\n\nSplice at: {1:F3}s\nVolume duck: {2:F1} dB\nOffset: {3}%\nFade type: {4}",
            successCount, spliceTime, duckDb, (int)(offsetPercent * 100), fadeType);
        if (errorCount > 0)
        {
            message += String.Format("\n\n{0} clip(s) had errors:\n{1}", errorCount, lastError);
        }
        MessageBox.Show(message, "Chorus Crisp Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ProcessEvent(TrackEvent originalEvent, double spliceTime, 
        double duckDb, double offsetPercent, CurveType fadeType)
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

        Timecode overlapDuration = Timecode.FromSeconds(spliceTime * offsetPercent);

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

public class ChorusCrispSettings
{
    public int SpliceValue = 47;
    public int CrispValue = 100;
    public int OffsetValue = 0;
    public int CurveIndex = 4;
    
    private static string GetSettingsPath()
    {
        string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ChorusCrisp");
        if (!Directory.Exists(folder))
            Directory.CreateDirectory(folder);
        return Path.Combine(folder, "settings.txt");
    }
    
    public void Save()
    {
        try
        {
            string[] lines = new string[]
            {
                "SpliceValue=" + SpliceValue.ToString(),
                "CrispValue=" + CrispValue.ToString(),
                "OffsetValue=" + OffsetValue.ToString(),
                "CurveIndex=" + CurveIndex.ToString()
            };
            File.WriteAllLines(GetSettingsPath(), lines);
        }
        catch { }
    }
    
    public static ChorusCrispSettings Load()
    {
        ChorusCrispSettings settings = new ChorusCrispSettings();
        try
        {
            string path = GetSettingsPath();
            if (File.Exists(path))
            {
                string[] lines = File.ReadAllLines(path);
                foreach (string line in lines)
                {
                    string[] parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        int value;
                        if (Int32.TryParse(parts[1].Trim(), out value))
                        {
                            if (key == "SpliceValue") settings.SpliceValue = value;
                            else if (key == "CrispValue") settings.CrispValue = value;
                            else if (key == "OffsetValue") settings.OffsetValue = value;
                            else if (key == "CurveIndex") settings.CurveIndex = value;
                        }
                    }
                }
            }
        }
        catch { }
        return settings;
    }
}

public class ChorusCrispPreset
{
    public string Name;
    public int SpliceValue;
    public int CrispValue;
    public int OffsetValue;
    public int CurveIndex;
    
    public ChorusCrispPreset(string name, int splice, int crisp, int offset, int curve)
    {
        Name = name;
        SpliceValue = splice;
        CrispValue = crisp;
        OffsetValue = offset;
        CurveIndex = curve;
    }
    
    public override string ToString() { return Name; }
}

public class ChorusCrispDialog : Form
{
    private ComboBox presetCombo;
    private TrackBar spliceSlider;
    private TrackBar crispSlider;
    private TrackBar offsetSlider;
    private ComboBox curveCombo;
    private Label spliceLabel;
    private Label crispLabel;
    private Label offsetLabel;
    private Button okButton;
    private Button cancelButton;
    
    private bool isLoadingPreset = false;
    private ChorusCrispPreset[] presets;

    public int SplicePosition { get { return spliceSlider.Value; } }
    public int Crispness { get { return 100 - crispSlider.Value; } }
    public int OffsetAmount { get { return 100 - offsetSlider.Value; } }
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
        InitializePresets();
        InitializeComponent();
        LoadSettings();
    }
    
    private void InitializePresets()
    {
        // Presets: Name, SpliceSlider, CrispSlider (inverted), OffsetSlider (inverted), CurveIndex
        // CurveIndex: 0=Linear, 1=Fast, 2=Slow, 3=Sharp, 4=Smooth
        
        // For CrispSlider: duckDb = (100 - sliderValue) / 100 * -15
        // -3 dB -> (100 - x)/100 * -15 = -3 -> x = 80
        // -5 dB -> (100 - x)/100 * -15 = -5 -> x = 67
        
        // For OffsetSlider: offset% = 100 - sliderValue
        // 80% -> sliderValue = 20
        // 90% -> sliderValue = 10
        // 95% -> sliderValue = 5
        
        presets = new ChorusCrispPreset[]
        {
            new ChorusCrispPreset("Custom", -1, -1, -1, -1),
            new ChorusCrispPreset("Jario Style", 50, 80, 20, 4),
            new ChorusCrispPreset("Standard", 40, 80, 10, 0),
            new ChorusCrispPreset("Snappy", 35, 67, 5, 1)
        };
    }

    private void InitializeComponent()
    {
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.AutoScaleDimensions = new SizeF(96F, 96F);
        
        this.Text = "Chorus Crisp";
        this.ClientSize = new Size(450, 420);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.FromArgb(45, 45, 48);
        this.ForeColor = Color.White;

        int yPos = 20;

        // Preset selector
        Label presetTitle = new Label();
        presetTitle.Text = "Preset:";
        presetTitle.Location = new Point(20, yPos);
        presetTitle.Size = new Size(60, 20);
        presetTitle.ForeColor = Color.White;
        this.Controls.Add(presetTitle);

        presetCombo = new ComboBox();
        presetCombo.Location = new Point(85, yPos - 3);
        presetCombo.Size = new Size(345, 28);
        presetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        presetCombo.Font = new Font(this.Font.FontFamily, 10);
        foreach (ChorusCrispPreset preset in presets)
        {
            presetCombo.Items.Add(preset);
        }
        presetCombo.SelectedIndex = 0;
        presetCombo.SelectedIndexChanged += PresetCombo_SelectedIndexChanged;
        this.Controls.Add(presetCombo);

        yPos += 40;

        // Splice Position column
        Label spliceTitle = new Label();
        spliceTitle.Text = "Splice Position";
        spliceTitle.Location = new Point(20, yPos);
        spliceTitle.Size = new Size(120, 20);
        spliceTitle.ForeColor = Color.White;
        spliceTitle.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(spliceTitle);

        // Crispness column
        Label crispTitle = new Label();
        crispTitle.Text = "Volume Duck";
        crispTitle.Location = new Point(160, yPos);
        crispTitle.Size = new Size(120, 20);
        crispTitle.ForeColor = Color.White;
        crispTitle.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(crispTitle);

        // Offset column
        Label offsetTitle = new Label();
        offsetTitle.Text = "Offset";
        offsetTitle.Location = new Point(300, yPos);
        offsetTitle.Size = new Size(120, 20);
        offsetTitle.ForeColor = Color.White;
        offsetTitle.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(offsetTitle);

        yPos += 25;

        // Splice slider (vertical)
        spliceSlider = new TrackBar();
        spliceSlider.Orientation = Orientation.Vertical;
        spliceSlider.Location = new Point(55, yPos);
        spliceSlider.Size = new Size(45, 120);
        spliceSlider.Minimum = 0;
        spliceSlider.Maximum = 100;
        spliceSlider.Value = 47;
        spliceSlider.TickFrequency = 10;
        spliceSlider.ValueChanged += Slider_ValueChanged;
        this.Controls.Add(spliceSlider);

        // Crisp slider (vertical)
        crispSlider = new TrackBar();
        crispSlider.Orientation = Orientation.Vertical;
        crispSlider.Location = new Point(195, yPos);
        crispSlider.Size = new Size(45, 120);
        crispSlider.Minimum = 0;
        crispSlider.Maximum = 100;
        crispSlider.Value = 100;
        crispSlider.TickFrequency = 10;
        crispSlider.ValueChanged += Slider_ValueChanged;
        this.Controls.Add(crispSlider);

        // Offset slider (vertical)
        offsetSlider = new TrackBar();
        offsetSlider.Orientation = Orientation.Vertical;
        offsetSlider.Location = new Point(335, yPos);
        offsetSlider.Size = new Size(45, 120);
        offsetSlider.Minimum = 0;
        offsetSlider.Maximum = 100;
        offsetSlider.Value = 0;
        offsetSlider.TickFrequency = 10;
        offsetSlider.ValueChanged += Slider_ValueChanged;
        this.Controls.Add(offsetSlider);

        yPos += 125;

        // Splice value label
        spliceLabel = new Label();
        spliceLabel.Location = new Point(20, yPos);
        spliceLabel.Size = new Size(120, 20);
        spliceLabel.ForeColor = Color.FromArgb(100, 200, 255);
        spliceLabel.Text = "47%";
        spliceLabel.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(spliceLabel);

        // Crisp value label
        crispLabel = new Label();
        crispLabel.Location = new Point(160, yPos);
        crispLabel.Size = new Size(120, 20);
        crispLabel.ForeColor = Color.FromArgb(100, 200, 255);
        crispLabel.Text = "0.0 dB";
        crispLabel.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(crispLabel);

        // Offset value label
        offsetLabel = new Label();
        offsetLabel.Location = new Point(300, yPos);
        offsetLabel.Size = new Size(120, 20);
        offsetLabel.ForeColor = Color.FromArgb(100, 200, 255);
        offsetLabel.Text = "100%";
        offsetLabel.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(offsetLabel);

        yPos += 30;

        // Range labels
        Label spliceRange = new Label();
        spliceRange.Text = "(20ms - 60ms)";
        spliceRange.Location = new Point(20, yPos);
        spliceRange.Size = new Size(120, 16);
        spliceRange.ForeColor = Color.Gray;
        spliceRange.TextAlign = ContentAlignment.MiddleCenter;
        spliceRange.Font = new Font(this.Font.FontFamily, 8);
        this.Controls.Add(spliceRange);

        Label crispRange = new Label();
        crispRange.Text = "(0 to -15 dB)";
        crispRange.Location = new Point(160, yPos);
        crispRange.Size = new Size(120, 16);
        crispRange.ForeColor = Color.Gray;
        crispRange.TextAlign = ContentAlignment.MiddleCenter;
        crispRange.Font = new Font(this.Font.FontFamily, 8);
        this.Controls.Add(crispRange);

        Label offsetRange = new Label();
        offsetRange.Text = "(0% - 100%)";
        offsetRange.Location = new Point(300, yPos);
        offsetRange.Size = new Size(120, 16);
        offsetRange.ForeColor = Color.Gray;
        offsetRange.TextAlign = ContentAlignment.MiddleCenter;
        offsetRange.Font = new Font(this.Font.FontFamily, 8);
        this.Controls.Add(offsetRange);

        yPos += 25;

        Label curveTitle = new Label();
        curveTitle.Text = "Crossfade Curve:";
        curveTitle.Location = new Point(20, yPos);
        curveTitle.Size = new Size(130, 20);
        curveTitle.ForeColor = Color.White;
        this.Controls.Add(curveTitle);

        yPos += 25;

        curveCombo = new ComboBox();
        curveCombo.Location = new Point(20, yPos);
        curveCombo.Size = new Size(410, 28);
        curveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        curveCombo.Font = new Font(curveCombo.Font.FontFamily, 10);
        curveCombo.Items.AddRange(new string[] { "Linear", "Fast", "Slow", "Sharp", "Smooth" });
        curveCombo.SelectedIndex = 4;
        curveCombo.SelectedIndexChanged += CurveCombo_SelectedIndexChanged;
        this.Controls.Add(curveCombo);

        yPos += 45;

        okButton = new Button();
        okButton.Text = "Apply Crisp";
        okButton.Location = new Point(100, yPos);
        okButton.Size = new Size(120, 35);
        okButton.BackColor = Color.FromArgb(0, 122, 204);
        okButton.ForeColor = Color.White;
        okButton.FlatStyle = FlatStyle.Flat;
        okButton.DialogResult = DialogResult.OK;
        okButton.Click += OkButton_Click;
        this.Controls.Add(okButton);

        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(230, yPos);
        cancelButton.Size = new Size(120, 35);
        cancelButton.BackColor = Color.FromArgb(80, 80, 80);
        cancelButton.ForeColor = Color.White;
        cancelButton.FlatStyle = FlatStyle.Flat;
        cancelButton.DialogResult = DialogResult.Cancel;
        this.Controls.Add(cancelButton);

        this.AcceptButton = okButton;
        this.CancelButton = cancelButton;
    }
    
    private void LoadSettings()
    {
        isLoadingPreset = true;
        ChorusCrispSettings settings = ChorusCrispSettings.Load();
        spliceSlider.Value = Math.Max(0, Math.Min(100, settings.SpliceValue));
        crispSlider.Value = Math.Max(0, Math.Min(100, settings.CrispValue));
        offsetSlider.Value = Math.Max(0, Math.Min(100, settings.OffsetValue));
        curveCombo.SelectedIndex = Math.Max(0, Math.Min(4, settings.CurveIndex));
        UpdateLabels();
        isLoadingPreset = false;
        CheckForMatchingPreset();
    }
    
    private void SaveSettings()
    {
        ChorusCrispSettings settings = new ChorusCrispSettings();
        settings.SpliceValue = spliceSlider.Value;
        settings.CrispValue = crispSlider.Value;
        settings.OffsetValue = offsetSlider.Value;
        settings.CurveIndex = curveCombo.SelectedIndex;
        settings.Save();
    }
    
    private void PresetCombo_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (presetCombo.SelectedIndex <= 0) return;
        
        ChorusCrispPreset preset = presets[presetCombo.SelectedIndex];
        if (preset.SpliceValue < 0) return;
        
        isLoadingPreset = true;
        spliceSlider.Value = preset.SpliceValue;
        crispSlider.Value = preset.CrispValue;
        offsetSlider.Value = preset.OffsetValue;
        curveCombo.SelectedIndex = preset.CurveIndex;
        UpdateLabels();
        isLoadingPreset = false;
    }
    
    private void Slider_ValueChanged(object sender, EventArgs e)
    {
        UpdateLabels();
        if (!isLoadingPreset)
        {
            CheckForMatchingPreset();
        }
    }
    
    private void CurveCombo_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!isLoadingPreset)
        {
            CheckForMatchingPreset();
        }
    }
    
    private void CheckForMatchingPreset()
    {
        // Check if current values match any preset
        for (int i = 1; i < presets.Length; i++)
        {
            ChorusCrispPreset p = presets[i];
            if (spliceSlider.Value == p.SpliceValue &&
                crispSlider.Value == p.CrispValue &&
                offsetSlider.Value == p.OffsetValue &&
                curveCombo.SelectedIndex == p.CurveIndex)
            {
                isLoadingPreset = true;
                presetCombo.SelectedIndex = i;
                isLoadingPreset = false;
                return;
            }
        }
        // No match, set to Custom
        isLoadingPreset = true;
        presetCombo.SelectedIndex = 0;
        isLoadingPreset = false;
    }
    
    private void UpdateLabels()
    {
        spliceLabel.Text = String.Format("{0}%", spliceSlider.Value);
        double duckDb = 0.0 + ((100 - crispSlider.Value) / 100.0) * (-15.0 - 0.0);
        crispLabel.Text = String.Format("{0:F1} dB", duckDb);
        offsetLabel.Text = String.Format("{0}%", 100 - offsetSlider.Value);
    }
    
    private void OkButton_Click(object sender, EventArgs e)
    {
        SaveSettings();
    }
}
