namespace L5Xs
{
    internal class XMLfiles
    {
        public static string GetEmptyProgramXMLContents(string programName, string routineName, string controllerName, string softwareRevision)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <RSLogix5000Content SchemaRevision=""1.0"" SoftwareRevision=""" + softwareRevision + @""" TargetName=""" + programName + @""" TargetType =""Program"" ContainsContext=""true"" ExportDate=""" + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy") + @""" ExportOptions=""References NoRawData L5KData DecoratedData Context Dependencies ForceProtectedEncoding AllProjDocTrans"">
                        <Controller Use=""Context"" Name=""" + controllerName + @""">
                            <Programs Use=""Context"">
                                <Program Use=""Target"" Name=""" + programName + @""" TestEdits=""false"" MainRoutineName=""" + routineName + @""" Disabled=""false"" UseAsFolder=""false"">
                                    <Tags/>
                                    <Routines>
                                        <Routine Name=""" + routineName + @""" Type=""RLL""/>
                                    </Routines>
                                </Program>
                            </Programs>
                        </Controller>
                    </RSLogix5000Content>";
        }

        public static string GetTaskXMLContents(string taskName, string programName, string controllerName, string softwareRevision)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <RSLogix5000Content SchemaRevision=""1.0"" SoftwareRevision=""" + softwareRevision + @""" TargetName=""ImpExp"" TargetType=""Task"" ContainsContext=""true"" ExportDate=""" + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy") + @""" ExportOptions=""References NoRawData L5KData DecoratedData Context ProductDefinedTypes IOTags Dependencies ForceProtectedEncoding AllProjDocTrans"">
                        <Controller Use=""Context"" Name=""" + controllerName + @""">
                            <Programs Use=""Context"">
                                <Program Use=""Reference"" Name=""" + programName + @""">
                                </Program>
                            </Programs>
                            <Tasks Use=""Context"">
                                <Task Use=""Target"" Name=""" + taskName + @""" Type=""CONTINUOUS"" Priority=""10"" Watchdog=""500"" DisableUpdateOutputs=""false"" InhibitTask=""false"">
                                    <ScheduledPrograms>
                                        <ScheduledProgram Name=""" + programName + @"""/>
                                    </ScheduledPrograms>
                                </Task>
                            </Tasks>
                        </Controller>
                    </RSLogix5000Content>";
        }

        public static string GetFaultHandlingRungXMLContents(string controllerName)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <RSLogix5000Content SchemaRevision=""1.0"" SoftwareRevision=""36.00"" TargetType=""Rung"" TargetCount=""1"" ContainsContext=""true"" ExportDate=""" + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy") + @""" ExportOptions=""References NoRawData L5KData DecoratedData Context RoutineLabels AliasExtras IOTags NoStringData ForceProtectedEncoding AllProjDocTrans"">
                        <Controller Use=""Context"" Name=""" + controllerName + @""">
                            <DataTypes Use=""Context"">
                            </DataTypes>
                            <Tags Use=""Context"">
                                <Tag Name=""AT_ClearFault"" TagType=""Base"" DataType=""BOOL"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- clear the fault  type & code]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""BOOL"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                                <Tag Name=""AT_FaultCode"" TagType=""Base"" DataType=""DINT"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- contains fault code number]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                                <Tag Name=""AT_FaultType"" TagType=""Base"" DataType=""DINT"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- contains fault type number]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                            </Tags>
                            <Programs Use=""Context"">
                                <Program Use=""Context"" Name=""P00_AOI_Testing"">
                                    <Routines Use=""Context"">
                                        <Routine Use=""Context"" Name=""R00_AOI_Testing"">
                                            <RLLContent Use=""Context"">
                                                <Rung Use=""Target"" Number=""0"" Type=""N"">
                                                    <Comment>
                                                        <![CDATA[AUTOMATED TESTING | FAULT HANDLING
                                                        - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                                                        Clear the tags storing the fault Type and Code information.]]>
                                                    </Comment>
                                                    <Text>
                                                        <![CDATA[XIC(AT_ClearFault)CLR(AT_FaultType)CLR(AT_FaultCode);]]>
                                                    </Text>
                                                </Rung>
                                            </RLLContent>
                                        </Routine>
                                    </Routines>
                                </Program>
                            </Programs>
                        </Controller>
                    </RSLogix5000Content>";
        }

        public static string GetFaultHandlerProgramXMLContents(string controllerName)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <RSLogix5000Content SchemaRevision=""1.0"" SoftwareRevision=""36.00"" TargetName=""PXX_FaultHandler"" TargetType=""Program"" ContainsContext=""true"" ExportDate=""" + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy") + @""" ExportOptions=""References NoRawData L5KData DecoratedData Context Dependencies ForceProtectedEncoding AllProjDocTrans"">
                        <Controller Use=""Context"" Name=""" + controllerName + @""">
                            <DataTypes Use=""Context"">
                                <DataType Name=""FAULTRECORD"" Family=""NoFamily"" Class=""User"">
                                    <Members>
                                        <Member Name=""Time_Low"" DataType=""DINT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Lower 32 bits of the fault timestamp value]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Time_High"" DataType=""DINT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Upper 32 bits of the fault timestamp value]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Type"" DataType=""INT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Fault type (program, I/O, and so forth)]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Code"" DataType=""INT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Unique code for the fault]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Info"" DataType=""DINT"" Dimension=""8"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Fault specific information]]>
                                            </Description>
                                        </Member>
                                    </Members>
                                </DataType>
                            </DataTypes>
                            <Tags Use=""Context"">
                                <Tag Name=""AT_FaultCode"" TagType=""Base"" DataType=""DINT"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- contains fault code number]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                                <Tag Name=""AT_FaultedIndicator"" TagType=""Base"" DataType=""BOOL"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- boolean HIGH if controller faulted]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""BOOL"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                                <Tag Name=""AT_FaultType"" TagType=""Base"" DataType=""DINT"" Radix=""Decimal"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Description>
                                        <![CDATA[Automated Testing
                                        -------------------- contains fault type number]]>
                                    </Description>
                                    <Data Format=""L5K"">
                                        <![CDATA[0]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <DataValue DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                    </Data>
                                </Tag>
                                <Tag Name=""major_fault_record"" TagType=""Base"" DataType=""FAULTRECORD"" Constant=""false"" ExternalAccess=""Read/Write"" OpcUaAccess=""None"">
                                    <Data Format=""L5K"">
                                        <![CDATA[[0,0,0,0,[0,0,0,0,0,0,0,0]]]]>
                                    </Data>
                                    <Data Format=""Decorated"">
                                        <Structure DataType=""FAULTRECORD"">
                                            <DataValueMember Name=""Time_Low"" DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                            <DataValueMember Name=""Time_High"" DataType=""DINT"" Radix=""Decimal"" Value=""0""/>
                                            <DataValueMember Name=""Type"" DataType=""INT"" Radix=""Decimal"" Value=""0""/>
                                            <DataValueMember Name=""Code"" DataType=""INT"" Radix=""Decimal"" Value=""0""/>
                                            <ArrayMember Name=""Info"" DataType=""DINT"" Dimensions=""8"" Radix=""Decimal"">
                                                <Element Index=""[0]"" Value=""0""/>
                                                <Element Index=""[1]"" Value=""0""/>
                                                <Element Index=""[2]"" Value=""0""/>
                                                <Element Index=""[3]"" Value=""0""/>
                                                <Element Index=""[4]"" Value=""0""/>
                                                <Element Index=""[5]"" Value=""0""/>
                                                <Element Index=""[6]"" Value=""0""/>
                                                <Element Index=""[7]"" Value=""0""/>
                                            </ArrayMember>
                                        </Structure>
                                    </Data>
                                </Tag>
                            </Tags>
                            <Programs Use=""Context"">
                                <Program Use=""Target"" Name=""PXX_FaultHandler"" TestEdits=""false"" MainRoutineName=""RXX_FaultHandler"" Disabled=""false"" UseAsFolder=""false"">
                                    <Tags/>
                                    <Routines>
                                        <Routine Name=""RXX_FaultHandler"" Type=""RLL"">
                                            <RLLContent>
                                                <Rung Number=""0"" Type=""N"">
                                                    <Comment>
                                                        <![CDATA[FAULT HANDLER
                                                        - - - - - - - - - - - - - - -
                                                        Get the fault (GSV) and store the fault Type and Code (MOVEs) in two tags that can be programmatically obtained during testing for better troubleshooting assistance.
                                                        Clear the fault Type and Code (SSV) using the major_fault_record tag for the FAULTRECORD user-defined data type.
                                                        Keep an indicator of fault state (OTE) for automated testing logic.

                                                        (Rung created with the assistance of https://literature.rockwellautomation.com/idc/groups/literature/documents/pm/1756-pm014_-en-p.pdf)
                                                        (Fault code spreadsheet link: https://literature.rockwellautomation.com/idc/groups/literature/documents/rd/1756-rd001_-en-p.xlsx)
                                                        ]]>
                                                    </Comment>
                                                    <Text>
                                                        <![CDATA[GSV(Program,THIS,MajorFaultRecord,major_fault_record.Time_Low)[MOVE(major_fault_record.Type,AT_FaultType) MOVE(major_fault_record.Code,AT_FaultCode) ,CLR(major_fault_record.Type) CLR(major_fault_record.Code) ,SSV(Program,THIS,MajorFaultRecord,major_fault_record.Time_Low) ,OTE(AT_FaultedIndicator) ];]]>
                                                    </Text>
                                                </Rung>
                                            </RLLContent>
                                        </Routine>
                                    </Routines>
                                </Program>
                            </Programs>
                        </Controller>
                    </RSLogix5000Content>";
        }


        public static string GetFaultRecordUDTXMLContents(string controllerName)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <RSLogix5000Content SchemaRevision=""1.0"" SoftwareRevision=""36.00"" TargetName=""FAULTRECORD"" TargetType=""DataType"" ContainsContext=""true"" ExportDate=""" + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy") + @""" ExportOptions=""References NoRawData L5KData DecoratedData Context Dependencies ForceProtectedEncoding AllProjDocTrans"">
                        <Controller Use=""Context"" Name=""" + controllerName + @""">
                            <DataTypes Use=""Context"">
                                <DataType Use=""Target"" Name=""FAULTRECORD"" Family=""NoFamily"" Class=""User"">
                                    <Members>
                                        <Member Name=""Time_Low"" DataType=""DINT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Lower 32 bits of the fault timestamp value]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Time_High"" DataType=""DINT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Upper 32 bits of the fault timestamp value]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Type"" DataType=""INT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Fault type (program, I/O, and so forth)]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Code"" DataType=""INT"" Dimension=""0"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Unique code for the fault]]>
                                            </Description>
                                        </Member>
                                        <Member Name=""Info"" DataType=""DINT"" Dimension=""8"" Radix=""Decimal"" Hidden=""false"" ExternalAccess=""Read/Write"">
                                            <Description>
                                                <![CDATA[Fault specific information]]>
                                            </Description>
                                        </Member>
                                    </Members>
                                </DataType>
                            </DataTypes>
                        </Controller>
                    </RSLogix5000Content>";
        }




    }
}
