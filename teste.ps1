Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Função para listar tarefas agendadas
function Listar-TarefasAgendadas {
    $taskService = New-Object -ComObject Schedule.Service
    $taskService.Connect()
    $rootFolder = $taskService.GetFolder('\')
    $tasks = $rootFolder.GetTasks(0)
    return $tasks | Where-Object { $_.Name -like 'atualiza*' }
}

function Atualizar-ListaTarefas {
    # Limpe a lista atual de tarefas
    $listBox.Items.Clear()

    # Carregue novamente as tarefas na lista
    $tarefas = Listar-TarefasAgendadas
    $listBox.Items.AddRange(($tarefas | ForEach-Object { $_.Name }))
}


# Função para obter o tipo de recorrência
function Get-TipoRecorrencia {
    param (
        [int]$TriggerType
    )

    switch ($TriggerType) {
        2 { return 'Diariamente' }
        3 {
            $taskService = New-Object -ComObject Schedule.Service
            $taskService.Connect()
            $rootFolder = $taskService.GetFolder('\')
            $tasks = $rootFolder.GetTasks(0)

            foreach ($task in $tasks) {
                $triggers = $task.Definition.Triggers
                foreach ($trigger in $triggers) {
                    if ($trigger.Type -eq 3) {
                        $daysOfWeek = @()
                        for ($i = 0; $i -lt 7; $i++) {
                            if ($trigger.DaysOfWeek -band [math]::Pow(2, $i)) {
                                $daysOfWeek += Get-NomeDiaSemana $i
                            }
                        }

                        $recorrencia = "Semanalmente"
                        if ($daysOfWeek.Count -gt 0) {
                            $recorrencia += "$([Environment]::NewLine)Dias Selecionados = $($daysOfWeek -join ', ')"
                        }

                        return $recorrencia
                    }
                }
            }

            return 'Recorrência semanal não encontrada'
        }
        4 { return 'Mensalmente' }
        1 { return 'Única vez' }
        default { return 'Desconhecido' }
    }
}

# Função para obter o nome do dia da semana
function Get-NomeDiaSemana {
    param (
        [int]$numeroDia
    )

    $diasDaSemana = @('Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado')
    return $diasDaSemana[$numeroDia]
}

# Função para exibir detalhes da tarefa
function Exibir-Detalhes {
    $tarefaSelecionada = $listBox.SelectedItem

    # Verificar se um item foi selecionado na lista
    if ($tarefaSelecionada -eq $null) {
        $textBox.Text = "Nenhuma tarefa selecionada."
        return
    }

    $taskService = New-Object -ComObject Schedule.Service
    $taskService.Connect()
    $rootFolder = $taskService.GetFolder("\")
    
    # Verificar se a tarefa com o nome selecionado existe
    $task = $rootFolder.GetTask($tarefaSelecionada)
    if ($task -eq $null) {
        $textBox.Text = "Tarefa não encontrada."
        return
    }

    $detalhes = "Nome da Tarefa: $($task.Name)`r`n"
    
    # Obtenção das informações de recorrência, data e hora
    $triggers = $task.Definition.Triggers
    if ($triggers.Count -gt 0) {
        foreach ($trigger in $triggers) {
            if ($trigger.Repetition -ne $null) {
                $tipoRecorrencia = Get-TipoRecorrencia $trigger.Type
                $detalhes += "Recorrência: $tipoRecorrencia`r`n"
                
                if ($trigger.StartBoundary -ne $null) {
                    $startBoundary = [DateTime]::Parse($trigger.StartBoundary)
                    $detalhes += "Data: $($startBoundary.ToString('dd/MM/yyyy'))`r`n"
                    $detalhes += "Hora: $($startBoundary.ToString('HH:mm'))`r`n"
                }
            }
        }
    } else {
        $detalhes += "Recorrência não especificada`r`n"
    }

    # Exibir os detalhes no formato desejado
    $textBox.Text = $detalhes
}

# Função para atualizar detalhes quando a seleção na lista muda
function Atualizar-Detalhes {
    Exibir-Detalhes
}

function Abrir-FormularioEdicao {
    param (
        [string]$nomeTarefa
    )

    # Criar um novo formulário para edição
    $formEdicao = New-Object Windows.Forms.Form
    $formEdicao.Size = New-Object Drawing.Size(450, 250)

        # Define o ícone do formulário
    $iconPath = ".\iconeinfoway.ico" # Substitua pelo caminho do seu ícone
    if (Test-Path $iconPath) {
        $icon = New-Object System.Drawing.Icon $iconPath
        $formEdicao.Icon = $icon
    } else {
        Write-Host "Ícone não encontrado."
    }

    # Título do formulário com o nome da tarefa
    $formEdicao.Text = "Tarefa Alterada: $nomeTarefa"

    # Desabilita o botão de maximizar
    $formEdicao.MaximizeBox = $false

        # Define uma imagem de fundo
    $imagePath = ".\tela_secundaria.png" # Substitua pelo caminho da sua imagem
    if (Test-Path $imagePath) {
        $backgroundImage = [System.Drawing.Image]::FromFile($imagePath)
        $formEdicao.BackgroundImage = $backgroundImage
        $formEdicao.BackgroundImageLayout = "Stretch"  # Ajuste o estilo de layout conforme necessário
    } else {
        Write-Host "Imagem de fundo não encontrada."
    }

    # Rótulo "Nova Recorrência"
    $labelRecorrencia = New-Object Windows.Forms.Label
    $labelRecorrencia.Text = "Nova Recorrência:"
    $labelRecorrencia.Size = New-Object Drawing.Size(105, 30)
    $labelRecorrencia.Location = New-Object Drawing.Point(110, 30)   
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $corHexadecimal = "#e35000"
    $cor = [System.Drawing.Color]::FromArgb(
        [System.Convert]::ToInt32($corHexadecimal.Substring(1), 16)
    )
    $labelRecorrencia.BackColor = $cor
    $labelRecorrencia.Font = $fonte
    $formEdicao.Controls.Add($labelRecorrencia)

    # ComboBox para selecionar a nova recorrência
    $comboBoxRecorrencia = New-Object Windows.Forms.ComboBox
    $comboBoxRecorrencia.Size = New-Object Drawing.Size(100, 20)
    $comboBoxRecorrencia.Location = New-Object Drawing.Point(215, 30)
    $comboBoxRecorrencia.Items.Add("Única vez")
    $comboBoxRecorrencia.Items.Add("Diariamente")
    $comboBoxRecorrencia.Items.Add("Semanalmente")
    $comboBoxRecorrencia.Add_SelectedIndexChanged({
        # Ao selecionar a recorrência semanal, mostrar os checkboxes dos dias da semana
        if ($comboBoxRecorrencia.SelectedItem -eq "Semanalmente") {
            $formEdicao.Size = New-Object Drawing.Size(450, 325)
            $labelDias.Visible = $true
            $checkBoxDomingo.Visible = $true
            $checkBoxSegunda.Visible = $true
            $checkBoxTerca.Visible = $true
            $checkBoxQuarta.Visible = $true
            $checkBoxQuinta.Visible = $true
            $checkBoxSexta.Visible = $true
            $checkBoxSabado.Visible = $true
            # Ajuste a posição dos controles de data e hora
            $labelData.Location = New-Object Drawing.Point(110, 150)
            $datePickerData.Location = New-Object Drawing.Point(215, 150)
            $labelHora.Location = New-Object Drawing.Point(110, 200)
            $timePickerHora.Location = New-Object Drawing.Point(215, 200)
            $buttonSalvar.Location = New-Object Drawing.Point(110, 240)
            $buttonCancelar.Location = New-Object Drawing.Point(215, 240)
        } else {
            $formEdicao.Size = New-Object Drawing.Size(450, 250)
            $labelDias.Visible = $false
            $checkBoxDomingo.Visible = $false
            $checkBoxSegunda.Visible = $false
            $checkBoxTerca.Visible = $false
            $checkBoxQuarta.Visible = $false
            $checkBoxQuinta.Visible = $false
            $checkBoxSexta.Visible = $false
            $checkBoxSabado.Visible = $false
            # Ajuste a posição dos controles de data e hora de volta ao padrão
            $labelData.Location = New-Object Drawing.Point(110, 75)
            $datePickerData.Location = New-Object Drawing.Point(215, 75)
            $labelHora.Location = New-Object Drawing.Point(110, 125)
            $timePickerHora.Location = New-Object Drawing.Point(215, 125)
            $buttonSalvar.Location = New-Object Drawing.Point(110, 170)
            $buttonCancelar.Location = New-Object Drawing.Point(215, 170)
        }
    })
    $formEdicao.Controls.Add($comboBoxRecorrencia)

    # Rótulo "Dias da Semana"
    $labelDias = New-Object Windows.Forms.Label
    $labelDias.Text = "Dias da Semana:"
    $labelDias.Size = New-Object Drawing.Size(100, 20)
    $labelDias.Location = New-Object Drawing.Point(50, 70)
    $labelDias.Visible = $false
    $labelDias.BackColor = [System.Drawing.Color]::Transparent
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labelDias.Font = $fonte
    $formEdicao.Controls.Add($labelDias)

    # Checkboxes para selecionar os dias da semana
    $checkBoxDomingo = New-Object Windows.Forms.CheckBox
    $checkBoxDomingo.Text = "Domingo"
    $checkBoxDomingo.Size = New-Object Drawing.Size(100, 20)
    $checkBoxDomingo.Location = New-Object Drawing.Point(150, 70)
    $checkBoxDomingo.Visible = $false
    $checkBoxDomingo.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxDomingo)

    $checkBoxSegunda = New-Object Windows.Forms.CheckBox
    $checkBoxSegunda.Text = "Segunda"
    $checkBoxSegunda.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSegunda.Location = New-Object Drawing.Point(250, 70)
    $checkBoxSegunda.Visible = $false
    $checkBoxSegunda.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSegunda)

    $checkBoxTerca = New-Object Windows.Forms.CheckBox
    $checkBoxTerca.Text = "Terça"
    $checkBoxTerca.Size = New-Object Drawing.Size(100, 20)
    $checkBoxTerca.Location = New-Object Drawing.Point(350, 70)
    $checkBoxTerca.Visible = $false
    $checkBoxTerca.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxTerca)

    $checkBoxQuarta = New-Object Windows.Forms.CheckBox
    $checkBoxQuarta.Text = "Quarta"
    $checkBoxQuarta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxQuarta.Location = New-Object Drawing.Point(50, 100)
    $checkBoxQuarta.Visible = $false
    $checkBoxQuarta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxQuarta)

    $checkBoxQuinta = New-Object Windows.Forms.CheckBox
    $checkBoxQuinta.Text = "Quinta"
    $checkBoxQuinta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxQuinta.Location = New-Object Drawing.Point(150, 100)
    $checkBoxQuinta.Visible = $false
    $checkBoxQuinta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxQuinta)

    $checkBoxSexta = New-Object Windows.Forms.CheckBox
    $checkBoxSexta.Text = "Sexta"
    $checkBoxSexta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSexta.Location = New-Object Drawing.Point(250, 100)
    $checkBoxSexta.Visible = $false
    $checkBoxSexta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSexta)

    $checkBoxSabado = New-Object Windows.Forms.CheckBox
    $checkBoxSabado.Text = "Sábado"
    $checkBoxSabado.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSabado.Location = New-Object Drawing.Point(350, 100)
    $checkBoxSabado.Visible = $false
    $checkBoxSabado.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSabado)

    # Rótulo "Nova Data"
    $labelData = New-Object Windows.Forms.Label
    $labelData.Text = "Nova Data:"
    $labelData.Size = New-Object Drawing.Size(80, 20)
    $labelData.Location = New-Object Drawing.Point(110, 75)
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $corHexadecimal = "#e35000"
    $cor = [System.Drawing.Color]::FromArgb(
        [System.Convert]::ToInt32($corHexadecimal.Substring(1), 16)
    )
    $labelData.BackColor = $cor
    $labelData.Font = $fonte    
    $formEdicao.Controls.Add($labelData)

    # Controle para escolher a nova data
    $datePickerData = New-Object Windows.Forms.DateTimePicker
    $datePickerData.Format = [Windows.Forms.DateTimePickerFormat]::Short
    $datePickerData.Size = New-Object Drawing.Size(100, 30)
    $datePickerData.Location = New-Object Drawing.Point(215, 75)
    $formEdicao.Controls.Add($datePickerData)

    # Rótulo "Novo Horário"
    $labelHora = New-Object Windows.Forms.Label
    $labelHora.Text = "Novo Horário:"
    $labelHora.Size = New-Object Drawing.Size(80, 30)
    $labelHora.Location = New-Object Drawing.Point(110, 125)
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $corHexadecimal = "#e35000"
    $cor = [System.Drawing.Color]::FromArgb(
        [System.Convert]::ToInt32($corHexadecimal.Substring(1), 16)
    )
    $labelHora.BackColor = $cor
    $labelHora.Font = $fonte    
    $formEdicao.Controls.Add($labelHora)

    # Controle para escolher o novo horário
    $timePickerHora = New-Object Windows.Forms.DateTimePicker
    $timePickerHora.Format = [Windows.Forms.DateTimePickerFormat]::Time
    $timePickerHora.ShowUpDown = $true
    $timePickerHora.Size = New-Object Drawing.Size(100, 30)
    $timePickerHora.Location = New-Object Drawing.Point(215, 125)
    $formEdicao.Controls.Add($timePickerHora)

        # Botão "Salvar"
    $buttonSalvar = New-Object Windows.Forms.Button
    $buttonSalvar.Text = "Salvar"
    $buttonSalvar.Size = New-Object Drawing.Size(100, 30)
    $buttonSalvar.Location = New-Object Drawing.Point(110, 170)
    $buttonSalvar.Add_Click({
        # Verificar se a recorrência foi selecionada
        if ($comboBoxRecorrencia.SelectedItem -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione uma recorrência.", "Erro")
            return
        }

        # Verificar se a nova data foi selecionada
        if ($datePickerData.Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione uma data.", "Erro")
            return
        }

        # Verificar se o novo horário foi selecionado
        if ($timePickerHora.Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione um horário.", "Erro")
            return
        }

        # Lógica para salvar as informações editadas da tarefa aqui
        $novaRecorrencia = $comboBoxRecorrencia.SelectedItem.ToString()
        $novaData = $datePickerData.Value
        $novoHorario = $timePickerHora.Value

        # Combine a data e a hora em uma única data
        $novaDataHora = $novaData.Date.Add($novoHorario.TimeOfDay)

        # Lógica para modificar a tarefa aqui
        # Utilize a função New-ScheduledTaskTrigger para criar um novo gatilho com base nas informações fornecidas
        # Utilize New-ScheduledTaskAction para definir a ação da tarefa
        # Utilize Set-ScheduledTask para atualizar a tarefa agendada existente

        # Exemplo de configuração de recorrência (substitua com a lógica real):
        switch ($novaRecorrencia) {
            "Única vez" {
                $trigger = New-ScheduledTaskTrigger -Once -At $novaDataHora
            }
            "Diariamente" {
                $trigger = New-ScheduledTaskTrigger -Daily -At $novoHorario
            }
            # ...
            "Semanalmente" {
                # Configura os dias da semana com base nos checkboxes selecionados
                $selectedDays = @()
                if ($checkBoxDomingo.Checked) {
                    $selectedDays += [System.DayOfWeek]::Sunday
                }
                if ($checkBoxSegunda.Checked) {
                    $selectedDays += [System.DayOfWeek]::Monday
                }
                if ($checkBoxTerca.Checked) {
                    $selectedDays += [System.DayOfWeek]::Tuesday
                }
                if ($checkBoxQuarta.Checked) {
                    $selectedDays += [System.DayOfWeek]::Wednesday
                }
                if ($checkBoxQuinta.Checked) {
                    $selectedDays += [System.DayOfWeek]::Thursday
                }
                if ($checkBoxSexta.Checked) {
                    $selectedDays += [System.DayOfWeek]::Friday
                }
                if ($checkBoxSabado.Checked) {
                    $selectedDays += [System.DayOfWeek]::Saturday
                }

                # Certifique-se de que pelo menos um dia da semana foi selecionado
                if ($selectedDays.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("Selecione pelo menos um dia da semana.", "Erro")
                    return
                }

                # Configure o gatilho semanal com os dias da semana selecionados
                $trigger = New-ScheduledTaskTrigger -Weekly -At $novaDataHora -DaysOfWeek $selectedDays
            }
            default {
                [System.Windows.Forms.MessageBox]::Show("Opção de recorrência inválida.", "Erro")
                return
            }
        }

        # Configure a ação da tarefa (substitua com a lógica real)
        $action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -Command Get-Date'

        # Atualize a tarefa agendada existente (substitua "$nomeTarefa" com o nome real da tarefa)
        Set-ScheduledTask -TaskName "$nomeTarefa" -Trigger $trigger -Action $action

        [System.Windows.Forms.MessageBox]::Show("Tarefas modificadas com sucesso.", "Sucesso")

        Atualizar-ListaTarefas

        $formEdicao.Close()
    })
$formEdicao.Controls.Add($buttonSalvar)


        # Botão "Cancelar"
        $buttonCancelar = New-Object Windows.Forms.Button
        $buttonCancelar.Text = "Cancelar"
        $buttonCancelar.Size = New-Object Drawing.Size(100, 30)
        $buttonCancelar.Location = New-Object Drawing.Point(215, 170)
        $buttonCancelar.Add_Click({
            # Fechar o formulário de edição ao clicar em "Cancelar"
            Atualizar-ListaTarefas
            $formEdicao.Close()
        })
        $formEdicao.Controls.Add($buttonCancelar)

        # Exibir o formulário de edição
        $formEdicao.ShowDialog()
    }

function criar_tarefa {
        # Importa o módulo ScheduledTasks
    Import-Module ScheduledTasks

    # Criar um novo formulário para edição
    $formEdicao = New-Object Windows.Forms.Form
    $formEdicao.Size = New-Object Drawing.Size(450, 250)

        # Define o ícone do formulário
    $iconPath = ".\iconeinfoway.ico" # Substitua pelo caminho do seu ícone
    if (Test-Path $iconPath) {
        $icon = New-Object System.Drawing.Icon $iconPath
        $formEdicao.Icon = $icon
    } else {
        Write-Host "Ícone não encontrado."
    }

    # Título do formulário com o nome da tarefa
    $formEdicao.Text = "Criação de Tarefa"

    # Desabilita o botão de maximizar
    $formEdicao.MaximizeBox = $false

       # Define uma imagem de fundo
    $imagePath = ".\tela_secundaria.png" # Substitua pelo caminho da sua imagem
    if (Test-Path $imagePath) {
        $backgroundImage = [System.Drawing.Image]::FromFile($imagePath)
        $formEdicao.BackgroundImage = $backgroundImage
        $formEdicao.BackgroundImageLayout = "Stretch"  # Ajuste o estilo de layout conforme necessário
    } else {
        Write-Host "Imagem de fundo não encontrada."
    }

    # Label e ComboBox para selecionar o tipo de tarefa
    $labelTipoTarefa = New-Object Windows.Forms.Label
    $labelTipoTarefa.Text = "Nova Tarefa:"
    $labelTipoTarefa.Size = New-Object Drawing.Size(105, 30)
    $labelTipoTarefa.Location = New-Object Drawing.Point(110, 20)
    $labelTipoTarefa.BackColor = [System.Drawing.Color]::Transparent
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labelTipoTarefa.Font = $fonte
    $formEdicao.Controls.Add($labelTipoTarefa)

    $comboBoxTipoTarefa = New-Object Windows.Forms.ComboBox
    $comboBoxTipoTarefa.Size = New-Object Drawing.Size(100, 20)
    $comboBoxTipoTarefa.Location = New-Object Drawing.Point(215, 20)
    $comboBoxTipoTarefa.Items.Add("Practice")    
    $comboBoxTipoTarefa.Items.Add("Sucessor")
    $comboBoxTipoTarefa.Items.Add("Suprema")
    $formEdicao.Controls.Add($comboBoxTipoTarefa)

    # Rótulo "Nova Recorrência"
    $labelRecorrencia = New-Object Windows.Forms.Label
    $labelRecorrencia.Text = "Nova Recorrência:"
    $labelRecorrencia.Size = New-Object Drawing.Size(105, 30)
    $labelRecorrencia.Location = New-Object Drawing.Point(110, 60)   
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labelRecorrencia.BackColor = [System.Drawing.Color]::Transparent
    $labelRecorrencia.Font = $fonte
    $formEdicao.Controls.Add($labelRecorrencia)

    # ComboBox para selecionar a nova recorrência
    $comboBoxRecorrencia = New-Object Windows.Forms.ComboBox
    $comboBoxRecorrencia.Size = New-Object Drawing.Size(100, 20)
    $comboBoxRecorrencia.Location = New-Object Drawing.Point(215, 60)
    $comboBoxRecorrencia.Items.Add("Única vez")
    $comboBoxRecorrencia.Items.Add("Diariamente")
    $comboBoxRecorrencia.Items.Add("Semanalmente")
    $comboBoxRecorrencia.Add_SelectedIndexChanged({
        # Ao selecionar a recorrência semanal, mostrar os checkboxes dos dias da semana
        if ($comboBoxRecorrencia.SelectedItem -eq "Semanalmente") {
            $formEdicao.Size = New-Object Drawing.Size(450, 325)
            $labelDias.Visible = $true
            $checkBoxDomingo.Visible = $true
            $checkBoxSegunda.Visible = $true
            $checkBoxTerca.Visible = $true
            $checkBoxQuarta.Visible = $true
            $checkBoxQuinta.Visible = $true
            $checkBoxSexta.Visible = $true
            $checkBoxSabado.Visible = $true
            # Ajuste a posição dos controles de data e hora
            $labelData.Location = New-Object Drawing.Point(110, 160)
            $datePickerData.Location = New-Object Drawing.Point(215, 160)
            $labelHora.Location = New-Object Drawing.Point(110, 200)
            $timePickerHora.Location = New-Object Drawing.Point(215, 200)
            $buttonCancelar.Location = New-Object Drawing.Point(215, 240)
            $buttonCriarTarefa.Location = New-Object Drawing.Point(110, 240)
          
        } else {
            $formEdicao.Size = New-Object Drawing.Size(450, 250)
            $labelDias.Visible = $false
            $checkBoxDomingo.Visible = $false
            $checkBoxSegunda.Visible = $false
            $checkBoxTerca.Visible = $false
            $checkBoxQuarta.Visible = $false
            $checkBoxQuinta.Visible = $false
            $checkBoxSexta.Visible = $false
            $checkBoxSabado.Visible = $false
            # Ajuste a posição dos controles de data e hora de volta ao padrão
            $labelData.Location = New-Object Drawing.Point(110, 95)
            $datePickerData.Location = New-Object Drawing.Point(215, 95)
            $labelHora.Location = New-Object Drawing.Point(110, 130)
            $timePickerHora.Location = New-Object Drawing.Point(215, 130)
            $buttonCancelar.Location = New-Object Drawing.Point(215, 170)
            $buttonCriarTarefa.Location = New-Object Drawing.Point(110, 170)
         
        }
    })
    $formEdicao.Controls.Add($comboBoxRecorrencia)

    # Rótulo "Dias da Semana"
    $labelDias = New-Object Windows.Forms.Label
    $labelDias.Text = "Dias da Semana:"
    $labelDias.Size = New-Object Drawing.Size(100, 20)
    $labelDias.Location = New-Object Drawing.Point(50, 100)
    $labelDias.Visible = $false
    $labelDias.BackColor = [System.Drawing.Color]::Transparent
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)    
    $labelDias.Font = $fonte
    $formEdicao.Controls.Add($labelDias)

    # Checkboxes para selecionar os dias da semana
    $checkBoxDomingo = New-Object Windows.Forms.CheckBox
    $checkBoxDomingo.Text = "Domingo"
    $checkBoxDomingo.Size = New-Object Drawing.Size(100, 20)
    $checkBoxDomingo.Location = New-Object Drawing.Point(150, 100)
    $checkBoxDomingo.Visible = $false
    $checkBoxDomingo.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxDomingo)

    $checkBoxSegunda = New-Object Windows.Forms.CheckBox
    $checkBoxSegunda.Text = "Segunda"
    $checkBoxSegunda.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSegunda.Location = New-Object Drawing.Point(250, 100)
    $checkBoxSegunda.Visible = $false
    $checkBoxSegunda.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSegunda)

    $checkBoxTerca = New-Object Windows.Forms.CheckBox
    $checkBoxTerca.Text = "Terça"
    $checkBoxTerca.Size = New-Object Drawing.Size(100, 20)
    $checkBoxTerca.Location = New-Object Drawing.Point(350, 100)
    $checkBoxTerca.Visible = $false
    $checkBoxTerca.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxTerca)

    $checkBoxQuarta = New-Object Windows.Forms.CheckBox
    $checkBoxQuarta.Text = "Quarta"
    $checkBoxQuarta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxQuarta.Location = New-Object Drawing.Point(50, 120)
    $checkBoxQuarta.Visible = $false
    $checkBoxQuarta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxQuarta)

    $checkBoxQuinta = New-Object Windows.Forms.CheckBox
    $checkBoxQuinta.Text = "Quinta"
    $checkBoxQuinta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxQuinta.Location = New-Object Drawing.Point(150, 120)
    $checkBoxQuinta.Visible = $false
    $checkBoxQuinta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxQuinta)

    $checkBoxSexta = New-Object Windows.Forms.CheckBox
    $checkBoxSexta.Text = "Sexta"
    $checkBoxSexta.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSexta.Location = New-Object Drawing.Point(250, 120)
    $checkBoxSexta.Visible = $false
    $checkBoxSexta.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSexta)

    $checkBoxSabado = New-Object Windows.Forms.CheckBox
    $checkBoxSabado.Text = "Sábado"
    $checkBoxSabado.Size = New-Object Drawing.Size(100, 20)
    $checkBoxSabado.Location = New-Object Drawing.Point(350, 120)
    $checkBoxSabado.Visible = $false
    $checkBoxSabado.BackColor = [System.Drawing.Color]::Transparent
    $formEdicao.Controls.Add($checkBoxSabado)

    # Rótulo "Nova Data"
    $labelData = New-Object Windows.Forms.Label
    $labelData.Text = "Nova Data:"
    $labelData.Size = New-Object Drawing.Size(80, 20)
    $labelData.Location = New-Object Drawing.Point(110, 95)
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)    
    $labelData.BackColor = [System.Drawing.Color]::Transparent
    $labelData.Font = $fonte    
    $formEdicao.Controls.Add($labelData)

    # Controle para escolher a nova data
    $datePickerData = New-Object Windows.Forms.DateTimePicker
    $datePickerData.Format = [Windows.Forms.DateTimePickerFormat]::Short
    $datePickerData.Size = New-Object Drawing.Size(100, 30)
    $datePickerData.Location = New-Object Drawing.Point(215, 95)
    $formEdicao.Controls.Add($datePickerData)

    # Rótulo "Novo Horário"
    $labelHora = New-Object Windows.Forms.Label
    $labelHora.Text = "Novo Horário:"
    $labelHora.Size = New-Object Drawing.Size(80, 30)
    $labelHora.Location = New-Object Drawing.Point(110, 130)
    $fonte = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labelHora.BackColor = [System.Drawing.Color]::Transparent
    $labelHora.Font = $fonte    
    $formEdicao.Controls.Add($labelHora)

    # Controle para escolher o novo horário
    $timePickerHora = New-Object Windows.Forms.DateTimePicker
    $timePickerHora.Format = [Windows.Forms.DateTimePickerFormat]::Time
    $timePickerHora.ShowUpDown = $true
    $timePickerHora.Size = New-Object Drawing.Size(100, 30)
    $timePickerHora.Location = New-Object Drawing.Point(215, 130)
    $formEdicao.Controls.Add($timePickerHora)       
    
    # Botão para cancelar
    $buttonCancelar = New-Object Windows.Forms.Button
    $buttonCancelar.Text = "Cancelar"
    $buttonCancelar.Size = New-Object Drawing.Size(100, 30)
    $buttonCancelar.Location = New-Object Drawing.Point(215, 170)    
    $buttonCancelar.Add_Click({
        # Fecha o formulário
        Atualizar-ListaTarefas
        $formEdicao.Close()
    })
    

        # Botão para criar a tarefa
    $buttonCriarTarefa = New-Object Windows.Forms.Button
    $buttonCriarTarefa.Text = "Criar Tarefa"
    $buttonCriarTarefa.Size = New-Object Drawing.Size(100, 30)
    $buttonCriarTarefa.Location = New-Object Drawing.Point(110, 170)

    $buttonCriarTarefa.Add_Click({

    # Verificar se a recorrência foi selecionada
        if ($comboBoxTipoTarefa.SelectedItem -eq $null){
        [System.Windows.Forms.MessageBox]::Show("Selecione um Tipo de Tarefa.", "Erro")
            return
        }
        if ($comboBoxRecorrencia.SelectedItem -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione uma recorrência.", "Erro")
            return
        }

        # Verificar se a nova data foi selecionada
        if ($datePickerData.Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione uma data.", "Erro")
            return
        }

        # Verificar se o novo horário foi selecionado
        if ($timePickerHora.Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Selecione um horário.", "Erro")
            return
        }

    # Coleta os valores selecionados e preenchidos nos controles do formulário
    $tipoTarefa = $comboBoxTipoTarefa.SelectedItem.ToString()
    $recorrencia = $comboBoxRecorrencia.SelectedItem.ToString()

    # Inicialize $taskname com um valor padrão, caso nenhum caso do switch seja correspondido
    $taskname = "Tarefa Padrão"

    # Verifica o tipo de tarefa selecionado e preenche os campos relacionados
    switch ($tipoTarefa) {
        "Practice" {
            $taskname = "Atualizacao Practice"
            $descricao = "Atualizacao do sistema Practice"
            $usuario = "Atualizacao"
            $caminhoAtualizacao = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-practice\executarscript.bat"
            $argumento = "executarscript.bat"
            $iniciarEm = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-practice"
        }
        "Sucessor" {
            $taskname = "Atualizacao Sucessor"
            $descricao = "Atualizacao do sistema Sucessor"
            $usuario = "Atualizacao"
            $caminhoAtualizacao = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-sucessor\executarscript.bat"
            $argumento = "executarscript.bat"
            $iniciarEm = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-sucessor"
        }
        "Suprema" {
            $taskname = "Atualizacao Suprema"
            $descricao = "Atualizacao do sistema Suprema"
            $usuario = "Atualizacao"
            $caminhoAtualizacao = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-suprema\executarscript.bat"
            $argumento = "executarscript.bat"
            $iniciarEm = "\\192.168.12.199\atualiza\Bot-atualizacao\dist-suprema"
        }
    }

    # Coleta a data e a hora selecionadas pelo usuário
    $dataSelecionada = $datePickerData.Value
    $horaSelecionada = $timePickerHora.Value

    # Verifica se o grupo "Usuários da área de trabalho remota" existe
    $grupoExiste = Get-LocalGroup -Name "Usuários da área de trabalho remota" -ErrorAction SilentlyContinue

    if (-not $grupoExiste) {
        # Cria o grupo "Usuários da área de trabalho remota" se ele não existir
        New-LocalGroup -Name "Usuários da área de trabalho remota" -Description "Grupo para Usuários da Área de Trabalho Remota"
        Write-Host "O grupo 'Usuários da área de trabalho remota' foi criado com sucesso."
    } else {
        Write-Host "O grupo 'Usuários da área de trabalho remota' já existe."
    }

    # Verifica se o usuário "Atualizacao" já existe
    $usuarioExiste = Get-LocalUser -Name "Atualizacao" -ErrorAction SilentlyContinue

    if (-not $usuarioExiste) {
        $password = "S@nta799"
        $objOu = [ADSI]"WinNT://localhost"

        try {
            $objUser = $objOu.Create("User", "Atualizacao")
            $objUser.SetPassword($password)
            $objUser.Put("Description", "Atualização Automática")
            $objUser.SetInfo()

            # Adiciona o usuário ao grupo "Usuários da área de trabalho remota"
            net localgroup "Usuários da área de trabalho remota" "Atualizacao" /add

            # Define a senha para nunca expirar
            wmic useraccount where "name='Atualizacao'" set passwordexpires=false

            Write-Host "Usuário 'Atualizacao' criado com sucesso."
            

           # Obtém o nome do domínio do ambiente
            $domain = $env:USERDOMAIN

            if (-not [string]::IsNullOrEmpty($domain)) {
                # Configura o Autologon
                $autologonPath = ".\Autologon.exe"  # Substitua pelo caminho correto
                $autologonUsername = "Atualizacao"
                $autologonPassword = "S@nta799"

                try {
                    # Define o usuário, domínio e senha para o Autologon
                    Start-Process -FilePath $autologonPath -ArgumentList $autologonUsername, $domain, $autologonPassword -Wait -NoNewWindow
                    Write-Host "Configuração de Autologon concluída com sucesso."
                } catch {
                    Write-Host "Erro ao configurar o Autologon: $_"
                }
            } else {
                Write-Host "Não foi possível determinar o domínio do usuário."
            }
        } catch {
            Write-Host "Erro ao criar o usuário 'Atualizacao': $_"
        }
    } else {
        # O usuário já existe, então apenas adicione-o ao grupo
        $usuario = 'Atualizacao'
        $usuarioGrupo = Get-LocalGroupMember -Group 'Usuários da área de trabalho remota' | Where-Object { $_.Name -eq $usuario }
        if ($usuarioGrupo = 'Usuários da área de trabalho remota'){
            Write-Host "Usuário 'Atualizacao' já existe e já está no grupo."
        }else{
            net localgroup "Usuários da área de trabalho remota" "Atualizacao" /add
            Write-Host "Usuário 'Atualizacao' já existe e foi adicionado ao grupo."
        }
    }


    $action = New-ScheduledTaskAction -Execute "$caminhoAtualizacao" -Argument "$argumento" -WorkingDirectory "$iniciarEm"

    # Cria um objeto de acionador para a recorrência selecionada
    switch ($recorrencia) {
        "Única vez" {
            $trigger = New-ScheduledTaskTrigger -Once -At $dataSelecionada
        }
        "Diariamente" {
            $trigger = New-ScheduledTaskTrigger -Daily -At $horaSelecionada
        }
        "Semanalmente" {
            $diasSelecionados = @()
            if ($checkBoxDomingo.Checked) { $diasSelecionados += "Sunday" }
            if ($checkBoxSegunda.Checked) { $diasSelecionados += "Monday" }
            if ($checkBoxTerca.Checked) { $diasSelecionados += "Tuesday" }
            if ($checkBoxQuarta.Checked) { $diasSelecionados += "Wednesday" }
            if ($checkBoxQuinta.Checked) { $diasSelecionados += "Thursday" }
            if ($checkBoxSexta.Checked) { $diasSelecionados += "Friday" }
            if ($checkBoxSabado.Checked) { $diasSelecionados += "Saturday" }

            $trigger = New-ScheduledTaskTrigger -Weekly -At $horaSelecionada -DaysOfWeek $diasSelecionados
        }        
    }
# Verifica se a tarefa já existe
if (Get-ScheduledTask -TaskName $taskname -ErrorAction SilentlyContinue) {
    [System.Windows.Forms.MessageBox]::Show("A tarefa já existe.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
} else {
    # Cria a tarefa agendada
    Register-ScheduledTask -TaskName $taskname -Action $action -Trigger $trigger -User $usuario -Description $descricao
    
    # Informa ao usuário que a tarefa foi criada com sucesso
    [System.Windows.Forms.MessageBox]::Show("Tarefa agendada criada com sucesso.", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    # Atualize a lista de tarefas após criar a tarefa
    Atualizar-ListaTarefas

    # Fecha o formulário
    $formEdicao.Close()
}})


# Adiciona o botão ao formulário
$formEdicao.Controls.Add($buttonCriarTarefa)
$formEdicao.Controls.Add($buttonCancelar)

# Exibir o formulário de edição
$formEdicao.ShowDialog()
}


# Cria uma instância do formulário principal
$form = New-Object Windows.Forms.Form
$form.Text = "Agendador de Tarefas"
$form.Size = New-Object Drawing.Size(460, 580)
$form.FormBorderStyle = "FixedSingle"  # Impede o redimensionamento


# Define o ícone do formulário
$iconPath = ".\iconeinfoway.ico" # Substitua pelo caminho do seu ícone
if (Test-Path $iconPath) {
    $icon = New-Object System.Drawing.Icon $iconPath
    $form.Icon = $icon
} else {
    Write-Host "Ícone não encontrado."
}

# Define uma imagem de fundo
$imagePath = ".\tela_principal.png" # Substitua pelo caminho da sua imagem
if (Test-Path $imagePath) {
    $backgroundImage = [System.Drawing.Image]::FromFile($imagePath)
    $form.BackgroundImage = $backgroundImage
    $form.BackgroundImageLayout = "Stretch"  # Ajuste o estilo de layout conforme necessário
} else {
    Write-Host "Imagem de fundo não encontrada."
}

# Desabilita o botão de maximizar
$form.MaximizeBox = $false

# Cria uma lista de tarefas (ListBox)
$listBox = New-Object Windows.Forms.ListBox
$listBox.Size = New-Object Drawing.Size(330, 200)
$listBox.Location = New-Object Drawing.Point(83, 67)
$form.Controls.Add($listBox)

# Cria uma caixa de texto para exibir detalhes
$textBox = New-Object Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.Size = New-Object Drawing.Size(330, 80)
$textBox.Location = New-Object Drawing.Point(83, 328)
$form.Controls.Add($textBox)

# Carrega as tarefas na lista
$tarefas = Listar-TarefasAgendadas
if ($tarefas -ne $null -and $tarefas.Count -gt 0) {
    $listBox.Items.AddRange(($tarefas | ForEach-Object { $_.Name }))
}

# Assine o evento SelectedIndexChanged na lista para atualizar detalhes automaticamente
$listBox.Add_SelectedIndexChanged({ Atualizar-Detalhes })

# Cria um botão "Alterar Horário"
$buttonAlterarHorario = New-Object Windows.Forms.Button
$buttonAlterarHorario.Text = "Alterar Horário"
$buttonAlterarHorario.Size = New-Object Drawing.Size(100, 30)
$buttonAlterarHorario.Location = New-Object Drawing.Point(98, 438)
$buttonAlterarHorario.Add_Click({
    # Verifica se uma tarefa está selecionada na lista
    $tarefaSelecionada = $listBox.SelectedItem
    if ($tarefaSelecionada -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Nenhuma tarefa selecionada.", "Erro")
        return
    }

    # Lógica para abrir o formulário de edição
    Abrir-FormularioEdicao -nomeTarefa $tarefaSelecionada.ToString()
})
$form.Controls.Add($buttonAlterarHorario)

# Cria um botão "Excluir"
$buttonExcluir = New-Object Windows.Forms.Button
$buttonExcluir.Text = "Excluir"
$buttonExcluir.Size = New-Object Drawing.Size(100, 30)
$buttonExcluir.Location = New-Object Drawing.Point(298, 438)
$buttonExcluir.Add_Click({
    # Verifica se uma tarefa está selecionada na lista
    $tarefaSelecionada = $listBox.SelectedItem
    if ($tarefaSelecionada -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Nenhuma tarefa selecionada.", "Erro")
        return
    }
    # Confirmação do usuário antes de excluir a tarefa
    $confirmacao = [System.Windows.Forms.MessageBox]::Show("Tem certeza de que deseja excluir a tarefa selecionada?", "Confirmação", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    
    if ($confirmacao -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            # Remove a tarefa agendada
            Unregister-ScheduledTask -TaskName $tarefaSelecionada.ToString() -Confirm:$false
            [System.Windows.Forms.MessageBox]::Show("Tarefa excluída com sucesso.", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Atualiza a lista de tarefas após a exclusão
            $tarefas = Listar-TarefasAgendadas
            $listBox.Items.Clear()
            if ($tarefas -ne $null -and $tarefas.Count -gt 0) {
                $listBox.Items.AddRange(($tarefas | ForEach-Object { $_.Name }))
            }
            Exibir-Detalhes
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao excluir a tarefa: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})
$form.Controls.Add($buttonExcluir)

# Cria um botão "Criar Tarefa"
$buttonCriar_tarefa = New-Object Windows.Forms.Button
$buttonCriar_tarefa.Text = "Criar Tarefa"
$buttonCriar_tarefa.Size = New-Object Drawing.Size(100, 30)
$buttonCriar_tarefa.Location = New-Object Drawing.Point(198, 438)
$buttonCriar_tarefa.Add_Click({
    criar_tarefa
})
$form.Controls.Add($buttonCriar_tarefa)


# Exibe o formulário principal
$form.ShowDialog()
