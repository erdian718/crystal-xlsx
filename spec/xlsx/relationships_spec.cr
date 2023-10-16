require "../spec_helper"

describe XLSX::Relationships do
  xml = <<-XML
    <Relationships>
      <Relationship Id="rId1" Type="type1" Target="target1"/>
      <Relationship Id="rId2" Type="type2" Target="target2"/>
      <Relationship Id="rId3" Type="type3" Target="target3"/>
    </Relationships>
  XML

  it "new" do
    rels = XLSX::Relationships.new(xml.to_slice)

    rels["rId1"].should eq("target1")
    rels["rId2"].should eq("target2")
    rels["rId3"].should eq("target3")
  end

  it "fetch" do
    rels = XLSX::Relationships.new(xml.to_slice)

    rels.fetch("type1").should eq("target1")
    rels.fetch("type2").should eq("target2")
    rels.fetch("type3").should eq("target3")

    target = rels.fetch("type4") do |type|
      type.should eq("type4")
      "target4"
    end
    target.should eq("target4")
    rels["rId4"].should eq("target4")
  end
end
