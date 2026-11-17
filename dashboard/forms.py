from django import forms


class ExcelUploadForm(forms.Form):
    file = forms.FileField(
        label="Planilha Excel",
        help_text="Envie um arquivo .xlsx: aceitamos o modelo simples ou a planilha completa do núcleo.",
    )

    def clean_file(self):
        uploaded_file = self.cleaned_data.get("file")
        if not uploaded_file:
            raise forms.ValidationError("Selecione um arquivo para importar.")

        extension = uploaded_file.name.lower().split(".")[-1]
        if extension != "xlsx":
            raise forms.ValidationError("Formato inválido. Use um arquivo .xlsx.")

        return uploaded_file
