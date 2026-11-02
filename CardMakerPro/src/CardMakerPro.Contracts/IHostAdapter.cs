using System;
using System.Collections.Generic;

namespace CardMakerPro.Contracts
{
    // Keep CorelDRAW interop out of the UI and shared engine.
    // Each supported CorelDRAW version supplies its own implementation.
    public interface IHostAdapter
    {
        string HostVersion { get; }
        bool HasActiveDocument { get; }
        void RunOnDocument(Action action);
        void SelectTargetShape();
        void PowerClip(string generatedShapeId, string targetShapeId, FitMode mode, Alignment alignment);
        void RotateShape(string shapeId, int degrees);
        RgbColor? PickNativeColor(RgbColor initialColor);
    }

    public enum FitMode { Contain, Cover, Stretch }
    public enum Alignment { Center, TopCenter, TopLeft, TopRight, BottomCenter }

    public struct RgbColor
    {
        public byte R, G, B;
        public RgbColor(byte r, byte g, byte b) { R = r; G = g; B = b; }
        public string Hex { get { return string.Format("#{0:X2}{1:X2}{2:X2}", R, G, B); } }
    }

    public enum BarcodeOrientation { Auto, Degrees0, Degrees90, Degrees180, Degrees270 }

    public sealed class BarcodeOptions
    {
        public BarcodeOrientation Orientation { get; set; }
        public bool PreserveQuietZone { get; set; }
        public bool ShowHumanReadableText { get; set; }
    }
}
