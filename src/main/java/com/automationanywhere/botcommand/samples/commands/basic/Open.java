/*
 * Copyright (c) 2020 Automation Anywhere.
 * All rights reserved.
 *
 * This software is the proprietary information of Automation Anywhere.
 * You shall use it only in accordance with the terms of the license agreement
 * you entered into with Automation Anywhere.
 */

/**
 *
 */
package com.automationanywhere.botcommand.samples.commands.basic;

import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;

import java.util.Map;

import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.Sessions;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;

/**
 * Sessions provide a mechanism to communicate between actions of the same
 * package. This means that sessions belonging to other packages are not visible
 * to action.
 * <p>
 * In this example we will check if a session( we will use string objects,
 * though in real actions a more complicated Object might be used) exists and if
 * not we will add it.
 *
 * @author Raj Singh Sisodia
 *
 */
//@BotCommand
@CommandPkg(
        label = "[[SessionDemo.label]]",
        description = "[[SessionDemo.description]]",
        icon = "sample.svg",
        name = "sessionDemo"
)
public class Open {

    // Sessions are provided as a Map. Actions can add or remove entries in this
    // Map.
    // The choice to reuse/overwrite/delete/add any Object in this Map belongs to
    // the actions, and the framework makes no assumption regarding it.
    @Sessions
    private Map<String, Object> sessionMap;

    @Execute
    public void execute(
            @Idx(index = "1", type = TEXT)
            @Pkg(label = "[[SessionDemo.name.label]]")
            @NotEmpty
                    String name
    ){
        if (!sessionMap.containsKey(name))
            sessionMap.put(name, "Some Value");
    }

    // Ensure that a public setter exists.
    public void setSessionMap(Map<String, Object> sessionMap) {
        this.sessionMap = sessionMap;
    }

}