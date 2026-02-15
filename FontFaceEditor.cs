using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Windows.Forms;

namespace ColorCallStack
{
    internal sealed class FontFaceEditor : UITypeEditor
    {
        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            return UITypeEditorEditStyle.Modal;
        }

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            string current = value as string;
            using (var dialog = new FontDialog
            {
                ShowColor = false,
                ShowEffects = false,
                AllowVectorFonts = true,
                AllowVerticalFonts = false,
                FontMustExist = true
            })
            {
                if (!string.IsNullOrWhiteSpace(current))
                {
                    try
                    {
                        dialog.Font = new Font(current, 10f);
                    }
                    catch
                    {
                        // Ignore invalid font names and show default selection.
                    }
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.Font?.Name ?? string.Empty;
                }
            }

            return value ?? string.Empty;
        }
    }
}
