# coding : utf8
import uno


def salaire_brut_chf(*args, **kwargs):
    contrat().salaire_brut_chf()

def salaire_brut_eur(*args, **kwargs):
    contrat().salaire_brut_eur()

def salaire_net_chf(*args, **kwargs):
    contrat().salaire_net_chf()

def salaire_net_eur(*args, **kwargs):
    contrat().salaire_net_eur()

def info_personnelles(*args, **kwargs):
    contrat().info_personnelles()

def toggle_euro(*args, **kwargs):
    contrat().toggle_euro()

class contrat():
    """
    Main class for everything tidied up
    """

    def __init__(self):
        self.document = XSCRIPTCONTEXT.getDocument()
        self.textfields = self.document.getTextFields()
        self.etranger = False
        self.chf_to_euro = 1.2
        self.list_tax = ('avs', 'ac', 'laa', 'cifav')
        self.value_dict = dict()
        self.tax_dict = dict()
        self.total_tax = 0

    def salaire_brut_chf(self):
        default = self.gettextfieldvalue('salaire')
        salaire = self.inputbox('salaire total',
                                'Entrez le salaire total brut !',
                                default)
        salaire = round(float(salaire), 1)
        self.calculate_brut_salary(salaire)


    def salaire_brut_eur(self):
        default = self.gettextfieldvalue('eur_salaire')
        euro_value = self.gettextfieldvalue('eur_val')
        salaire = self.inputbox('salaire total',
                                'Entrez le salaire total brut !',
                                default)
        salaire = round(float(salaire), 1) * euro_value
        self.calculate_brut_salary(salaire)

    def calculate_brut_salary(self, brut_salary):

        salaire = brut_salary
        taux_vacances = float(self.gettextfieldvalue('taux_vacances'))
        salaire_base_brut = round(salaire / (1 + taux_vacances), 1)
        vacances = salaire - round(salaire_base_brut, 1)

        self.get_tax_dict()

        for (tax, tax_value) in self.tax_dict.items():
            key = 'salaire_' + tax
            tax_price = round(salaire * tax_value, 1)
            self.total_tax += tax_price
            self.value_dict[key] = tax_price

        salaire_net = salaire - self.total_tax

        self.value_dict.update(
            {'salaire_brut': salaire,
             'salaire_base_brut': salaire_base_brut,
             'vacances': vacances,
             'total_tax': self.total_tax,
             'salaire_net': salaire_net,
             'salaire_net_2': salaire_net,
             }
        )

        self.update_all_fields(self.value_dict)

    def salaire_net_chf(self):
        default = self.gettextfieldvalue('salaire_net')
        salaire_net = self.inputbox('salaire total',
                                'Entrez le salaire total net !',
                                default)
        salaire_net = round(float(salaire_net), 1)
        self.calculate_net_salary(salaire_net)

    def salaire_net_eur(self):
        default = self.gettextfieldvalue('eur_salaire_net')
        eur_salaire_net = self.inputbox('salaire net en euro',
                                        'Entrez le salaire total',
                                        default)
        euro_value = self.gettextfieldvalue('eur_val')
        salaire_net = round(float(eur_salaire_net), 1) * euro_value
        self.calculate_net_salary(salaire_net)


    def calculate_net_salary(self, net_salary):
        salaire_net = float(net_salary)
        self.get_tax_dict()
        tax_sum = 0
        for (tax, tax_value) in self.tax_dict.items():
                tax_sum += tax_value
        salaire = round (salaire_net / (1 - tax_sum), 1)


        for (tax, tax_value) in self.tax_dict.items():
            key = 'salaire_' + tax
            tax_price = round(salaire * tax_value, 1)
            self.total_tax += tax_price
            self.value_dict[key] = tax_price
        salaire = salaire_net + self.total_tax

        taux_vacances = float(self.gettextfieldvalue('taux_vacances'))
        salaire_base_brut = round(salaire / (1 + taux_vacances), 1)
        vacances = salaire - salaire_base_brut

        self.value_dict.update(
            {'salaire_brut': salaire,
             'salaire_base_brut': salaire_base_brut,
             'vacances': vacances,
             'total_tax': self.total_tax,
             'salaire_net': salaire_net,
             'salaire_net_2': salaire_net,
             }
        )

        self.update_all_fields(self.value_dict)

    def info_personnelles(self):
        fields = ('rep_employeur', 'employeur', 'nomemployé', 'spectacle',
                  'qualité')
        for field in fields:
            self.update_dialog(field)

    def toggle_euro(self):
        toggled_fields_eur = (
            'is_text', 'isource', 'eur_salaire_isource', 'chf_salaire_isource',
            'eur_salaire_ac', 'eur_salaire_avs', 'eur_salaire_cifav',
            'eur_salaire_net', 'eur_salaire_laa', 'eur_total_tax',
            'eur_salaire_base_brut', 'eur_vacances', 'eur_salaire_brut',
            'eur_salaire_net_2',
            'chf_salaire_ac', 'chf_salaire_avs', 'chf_salaire_cifav',
            'chf_salaire_net', 'chf_salaire_laa', 'chf_total_tax',
            'chf_salaire_base_brut', 'chf_vacances', 'chf_salaire_brut',
            'chf_salaire_net_2',
            )

        toggled_fields_chf = (
            'salaire_ac', 'salaire_avs', 'salaire_cifav',
            'salaire_net', 'salaire_laa', 'total_tax',
            'salaire_base_brut', 'vacances', 'salaire_brut',
            'salaire_net_2',
            )
        euro_value = float(self.gettextfieldvalue('eur_val'))
        euro_bool = self.gettextfieldvalue('euro_bool')
        if euro_bool:
            self.updatetextfield('euro_bool', 0)
            for field in toggled_fields_eur:
                self.toggle_field(field, 0)
            for field in toggled_fields_chf:
                self.toggle_field(field, 1)
            net_salary = self.gettextfieldvalue('salaire_net')
            self.calculate_net_salary(net_salary)
        else:
            eur_chf = self.update_dialog('eur_val')
            self.updatetextfield('euro_bool', 1)
            for field in toggled_fields_eur:
                self.toggle_field(field, 1)
            for field in toggled_fields_chf:
                self.toggle_field(field, 0)
            self.list_tax = self.list_tax + ('isource', )
            net_salary = self.gettextfieldvalue('salaire_net')
            self.calculate_net_salary(net_salary)

        self.repercutate_eur_chf_prices()

    def repercutate_eur_chf_prices(self):

        euro_value = float(self.gettextfieldvalue('eur_val'))
        for (key, value) in self.value_dict.items():
            self.updatetextfield('chf_'+key, value)
            try:
                value = value/euro_value
            except:
                pass
            self.updatetextfield('eur_'+key, value)





    def toggle_field(self, title, value):
        enum = self.textfields.createEnumeration()
        while enum.hasMoreElements():
            tf = enum.nextElement()
            try:
                if tf.VariableName == title:
                    tf.IsVisible = value
                    tf.update()
            except AttributeError:
                pass

    def update_dialog(self, title):
        default = self.gettextfieldvalue(title)
        value = self.inputbox('Mettre à jour',
                              'mettre à jour le champ : " ' + title + ' "',
                              default)
        self.updatetextfield(title, value)
        return value

    def get_tax_dict(self):
        if 'isource' not in self.list_tax:
            eur = self.gettextfieldvalue('eur_bool')
            if eur == 1:
                self.list_tax = self.list_tax + ('isource', )
        for tax_name in self.list_tax:
            self.tax_dict[tax_name] = float(self.gettextfieldvalue(tax_name))

    def update_all_fields(self, value_dict):
        for key, val in self.value_dict.items():
            self.updatetextfield(key, val)

        self.repercutate_eur_chf_prices()

    def gettextfieldvalue(self, title):
        """find the field with title and return its value"""
        enum = self.textfields.createEnumeration()
        while enum.hasMoreElements():
            tf = enum.nextElement()
            try:
                if tf.VariableName == title:
                    try:
                        val = float(tf.getPropertyValue('Content'))
                    except ValueError:
                        val = tf.getPropertyValue('Content')
                    return val
            except AttributeError:
                pass

    def updatetextfield(self, title, value):
        """find the field (variable) with title and
        update its value to value"""
        enum = self.textfields.createEnumeration()
        while enum.hasMoreElements():
            tf = enum.nextElement()
            try:
                if tf.VariableName == title:
                    tf.Content = value
                    tf.update()
            except AttributeError:
                pass

    def inputbox(self, message, title="", default="", x=None, y=None):
        """ Shows dialog with input box.
            @param message message to show on the dialog
            @param title window title
            @param default default value
            @param x dialog positio in twips, pass y also
            @param y dialog position in twips, pass y also
            @return string if OK button pushed, otherwise zero length string
        """
        WIDTH = 300
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = VERT_SEP = 8
        LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
        EDIT_HEIGHT = 24
        HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
        import uno
        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
        from com.sun.star.awt.PushButtonType import OK, CANCEL
        from com.sun.star.util.MeasureUnit import TWIP
        ctx = uno.getComponentContext()

        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)

        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT, 
            {"Label": str(message), "NoLabel": True})
        add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        add("btn_cancel", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN + BUTTON_HEIGHT + 5, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})
        add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        edit = dialog.getControl("edit")
        edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
        edit.setFocus()
        ret = edit.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        try:
            ret = float(ret)
        except ValueError:
            ret = ret
        return ret
